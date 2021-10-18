using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 작번 비용 이관
	/// </summary>
	internal class PS_PP570 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP570L; //등록라인

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP570.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP570_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP570");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

				oForm.Freeze(true);

				PS_PP570_CreateItems();
				PS_PP570_SetComboBox();
				PS_PP570_Initialize();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", false); // 행삭제

				oForm.Items.Item("BOrdNum").Click();
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
			}
		}

		/// <summary>
		/// PS_PP570_CreateItems
		/// </summary>
		private void PS_PP570_CreateItems()
		{
			try
			{
				oDS_PS_PP570L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//이관전작번
				oForm.DataSources.UserDataSources.Add("BOrdNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 11);
				oForm.Items.Item("BOrdNum").Specific.DataBind.SetBound(true, "", "BOrdNum");

				//이관전서브1
				oForm.DataSources.UserDataSources.Add("BOrdSub1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
				oForm.Items.Item("BOrdSub1").Specific.DataBind.SetBound(true, "", "BOrdSub1");

				//이관전서브2
				oForm.DataSources.UserDataSources.Add("BOrdSub2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
				oForm.Items.Item("BOrdSub2").Specific.DataBind.SetBound(true, "", "BOrdSub2");

				//이관전품명
				oForm.DataSources.UserDataSources.Add("BItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("BItemName").Specific.DataBind.SetBound(true, "", "BItemName");

				//이관전규격
				oForm.DataSources.UserDataSources.Add("BItemSpec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("BItemSpec").Specific.DataBind.SetBound(true, "", "BItemSpec");

				//이관후작번
				oForm.DataSources.UserDataSources.Add("AOrdNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 11);
				oForm.Items.Item("AOrdNum").Specific.DataBind.SetBound(true, "", "AOrdNum");

				//이관후서브1
				oForm.DataSources.UserDataSources.Add("AOrdSub1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
				oForm.Items.Item("AOrdSub1").Specific.DataBind.SetBound(true, "", "AOrdSub1");

				//이관후서브2
				oForm.DataSources.UserDataSources.Add("AOrdSub2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
				oForm.Items.Item("AOrdSub2").Specific.DataBind.SetBound(true, "", "AOrdSub2");

				//이관후품명
				oForm.DataSources.UserDataSources.Add("AItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("AItemName").Specific.DataBind.SetBound(true, "", "AItemName");

				//이관후규격
				oForm.DataSources.UserDataSources.Add("AItemSpec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("AItemSpec").Specific.DataBind.SetBound(true, "", "AItemSpec");

				//등록자사번
				oForm.DataSources.UserDataSources.Add("TCntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("TCntcCode").Specific.DataBind.SetBound(true, "", "TCntcCode");

				//등록자성명
				oForm.DataSources.UserDataSources.Add("TCntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("TCntcName").Specific.DataBind.SetBound(true, "", "TCntcName");

				//요청자사번
				oForm.DataSources.UserDataSources.Add("RCntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("RCntcCode").Specific.DataBind.SetBound(true, "", "RCntcCode");

				//요청자성명
				oForm.DataSources.UserDataSources.Add("RCntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("RCntcName").Specific.DataBind.SetBound(true, "", "RCntcName");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP570_SetComboBox
		/// </summary>
		private void PS_PP570_SetComboBox()
		{
			try
			{
				//이관구분
				oForm.Items.Item("TransCls").Specific.ValidValues.Add("01", "전체비용");
				oForm.Items.Item("TransCls").Specific.ValidValues.Add("02", "공수");
				oForm.Items.Item("TransCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP570_Initialize
		/// </summary>
		private void PS_PP570_Initialize()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("TCntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP570_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP570_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP570L.InsertRecord(oRow);
				}

				oMat.AddRow();
				oDS_PS_PP570L.Offset = oRow;
				oDS_PS_PP570L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP570_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP570_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			try
			{
				oForm.Freeze(true);

				if (oUID == "TCntcCode")
				{
					sQry = " SELECT  U_FullName";
					sQry += " FROM    [@PH_PY001A]";
					sQry += " WHERE   Code = '" + oForm.Items.Item("TCntcCode").Specific.Value.ToString().Trim() + "'";
					oRecordSet.DoQuery(sQry);

					oForm.Items.Item("TCntcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
				}
				else if (oUID == "RCntcCode")
				{
					sQry = " SELECT  U_FullName";
					sQry += " FROM    [@PH_PY001A]";
					sQry += " WHERE   Code = '" + oForm.Items.Item("RCntcCode").Specific.Value.ToString().Trim() + "'";
					oRecordSet.DoQuery(sQry);

					oForm.Items.Item("RCntcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
				}
				else if (oUID == "BOrdNum")
				{
					sQry = " SELECT  U_JakMyung,";
					sQry += "         U_JakSize";
					sQry += " FROM    [@PS_PP030H]";
					sQry += " WHERE   U_OrdNum = '" + oForm.Items.Item("BOrdNum").Specific.Value.ToString().Trim() + "'";
					sQry += "         AND U_OrdSub1 = '00'";
					sQry += "         AND U_OrdSub2 = '000'";
					oRecordSet.DoQuery(sQry);

					oForm.Items.Item("BOrdSub1").Specific.Value = "00";
					oForm.Items.Item("BOrdSub2").Specific.Value = "000";
					oForm.Items.Item("BItemName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					oForm.Items.Item("BItemSpec").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
				}
				else if (oUID == "BOrdSub1")
				{
					sQry = " SELECT  U_JakMyung,";
					sQry += "         U_JakSize";
					sQry += " FROM    [@PS_PP030H]";
					sQry += " WHERE   U_OrdNum = '" + oForm.Items.Item("BOrdNum").Specific.Value.ToString().Trim() + "'";
					sQry += "         AND U_OrdSub1 = '" + oForm.Items.Item("BOrdSub1").Specific.Value.ToString().Trim() + "'";
					sQry += "         AND U_OrdSub2 = '000'";
					oRecordSet.DoQuery(sQry);

					oForm.Items.Item("BOrdSub2").Specific.Value = "000";
					oForm.Items.Item("BItemName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					oForm.Items.Item("BItemSpec").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
				}
				else if (oUID == "BOrdSub2")
				{
					sQry = " SELECT  U_JakMyung,";
					sQry += "         U_JakSize";
					sQry += " FROM    [@PS_PP030H]";
					sQry += " WHERE   U_OrdNum = '" + oForm.Items.Item("BOrdNum").Specific.Value.ToString().Trim() + "'";
					sQry += "         AND U_OrdSub1 = '" + oForm.Items.Item("BOrdSub1").Specific.Value.ToString().Trim() + "'";
					sQry += "         AND U_OrdSub2 = '" + oForm.Items.Item("BOrdSub2").Specific.Value.ToString().Trim() + "'";
					oRecordSet.DoQuery(sQry);

					oForm.Items.Item("BItemName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					oForm.Items.Item("BItemSpec").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
				}
				else if (oUID == "AOrdNum")
				{
					sQry = " SELECT  U_JakMyung,";
					sQry += "         U_JakSize";
					sQry += " FROM    [@PS_PP030H]";
					sQry += " WHERE   U_OrdNum = '" + oForm.Items.Item("AOrdNum").Specific.Value.ToString().Trim() + "'";
					sQry += "         AND U_OrdSub1 = '00'";
					sQry += "         AND U_OrdSub2 = '000'";
					oRecordSet.DoQuery(sQry);

					oForm.Items.Item("AOrdSub1").Specific.Value = "00";
					oForm.Items.Item("AOrdSub2").Specific.Value = "000";
					oForm.Items.Item("AItemName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					oForm.Items.Item("AItemSpec").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
				}
				else if (oUID == "AOrdSub1")
				{
					sQry = " SELECT  U_JakMyung,";
					sQry += "         U_JakSize";
					sQry += " FROM    [@PS_PP030H]";
					sQry += " WHERE   U_OrdNum = '" + oForm.Items.Item("AOrdNum").Specific.Value.ToString().Trim() + "'";
					sQry += "         AND U_OrdSub1 = '" + oForm.Items.Item("AOrdSub1").Specific.Value.ToString().Trim() + "'";
					sQry += "         AND U_OrdSub2 = '000'";
					oRecordSet.DoQuery(sQry);

					oForm.Items.Item("AOrdSub2").Specific.Value = "000";
					oForm.Items.Item("AItemName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					oForm.Items.Item("AItemSpec").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
				}
				else if (oUID == "AOrdSub2")
				{
					sQry = " SELECT  U_JakMyung,";
					sQry += "         U_JakSize";
					sQry += " FROM    [@PS_PP030H]";
					sQry += " WHERE   U_OrdNum = '" + oForm.Items.Item("AOrdNum").Specific.Value.ToString().Trim() + "'";
					sQry += "         AND U_OrdSub1 = '" + oForm.Items.Item("AOrdSub1").Specific.Value.ToString().Trim() + "'";
					sQry += "         AND U_OrdSub2 = '" + oForm.Items.Item("AOrdSub2").Specific.Value.ToString().Trim() + "'";
					oRecordSet.DoQuery(sQry);

					oForm.Items.Item("AItemName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					oForm.Items.Item("AItemSpec").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
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
		/// PS_PP570_CheckDataValid
		/// </summary>
		/// <returns></returns>
		private bool PS_PP570_CheckDataValid()
		{
			bool returnValue = false;
			short loopCount;
			string errMessage = string.Empty;

			try
			{
				//Header Check
				if (string.IsNullOrEmpty(oForm.Items.Item("TCntcCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "등록자정보는 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("RCntcCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "요청자정보는 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}

				//Line Check
				oMat.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					//체크 확인 로직 변경(FlushDataSource 후 에도 Check box 값을 못 가져옴, Matrix의 값을 직접 가져옴)
					if (oMat.Columns.Item("Check").Cells.Item(loopCount + 1).Specific.Checked == true)
					{
						if (string.IsNullOrEmpty(oDS_PS_PP570L.GetValue("U_ColReg18", loopCount).ToString().Trim()))
						{
							errMessage = "이관후 작번은 필수사항입니다. 확인하여 주십시오.";
							throw new Exception();
						}
						if (string.IsNullOrEmpty(oDS_PS_PP570L.GetValue("U_ColReg19", loopCount).ToString().Trim()))
						{
							errMessage = "이관후 서브작번1은 필수사항입니다. 확인하여 주십시오.";
							throw new Exception();
						}
						if (string.IsNullOrEmpty(oDS_PS_PP570L.GetValue("U_ColReg20", loopCount).ToString().Trim()))
						{
							errMessage = "이관후 서브작번2은 필수사항입니다. 확인하여 주십시오.";
							throw new Exception();
						}
					}
				}

				returnValue = true;
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
			return returnValue;
		}

		/// <summary>
		/// PS_PP570_MTX01 데이터 조회
		/// </summary>
		private void PS_PP570_MTX01()
		{
			short i;
			string sQry;
			string errMessage = string.Empty;
			string BOrdNum;  //이관전작번
			string BOrdSub1; //이관전서브1
			string BOrdSub2; //이관전서브2

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				ProgressBar01.Text = "조회시작!";

				oForm.Freeze(true);

				BOrdNum = oForm.Items.Item("BOrdNum").Specific.Value.ToString().Trim();
				BOrdSub1 = oForm.Items.Item("BOrdSub1").Specific.Value.ToString().Trim();
				BOrdSub2 = oForm.Items.Item("BOrdSub2").Specific.Value.ToString().Trim();

				sQry = "EXEC [PS_PP570_01] '";
				sQry += BOrdNum + "','";
				sQry += BOrdSub1 + "','";
				sQry += BOrdSub2 + "'";

				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_PP570L.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP570L.Size)
					{
						oDS_PS_PP570L.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_PP570L.Offset = i;
					oDS_PS_PP570L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP570L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Check").Value.ToString().Trim());	  //선택
					oDS_PS_PP570L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());	  //공정코드
					oDS_PS_PP570L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());	  //공정명
					oDS_PS_PP570L.SetValue("U_ColDt01", i, Convert.ToString(oRecordSet.Fields.Item("WorkDate").Value.ToString().Trim()).ToStrimg("yyyyMMdd")); //작업일자
					oDS_PS_PP570L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("WorkCode").Value.ToString().Trim());  //작업자사번
					oDS_PS_PP570L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("WorkName").Value.ToString().Trim());  //작업자성명
					oDS_PS_PP570L.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("WorkTime").Value.ToString().Trim());  //공수
					oDS_PS_PP570L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("PP040HNo").Value.ToString().Trim());  //작업일보번호
					oDS_PS_PP570L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("PP040LNo").Value.ToString().Trim());  //작업일보라인
					oDS_PS_PP570L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("BOrdNum").Value.ToString().Trim());	  //이관전작번
					oDS_PS_PP570L.SetValue("U_ColReg11", i, oRecordSet.Fields.Item("BOrdSub1").Value.ToString().Trim());  //이관전서브1
					oDS_PS_PP570L.SetValue("U_ColReg12", i, oRecordSet.Fields.Item("BOrdSub2").Value.ToString().Trim());  //이관전서브2
					oDS_PS_PP570L.SetValue("U_ColReg13", i, oRecordSet.Fields.Item("BItemName").Value.ToString().Trim()); //이관전품명
					oDS_PS_PP570L.SetValue("U_ColReg14", i, oRecordSet.Fields.Item("BItemSpec").Value.ToString().Trim()); //이관전규격
					oDS_PS_PP570L.SetValue("U_ColReg15", i, oRecordSet.Fields.Item("BPP030HNo").Value.ToString().Trim()); //이관전작지번호
					oDS_PS_PP570L.SetValue("U_ColReg16", i, oRecordSet.Fields.Item("BPP030MNo").Value.ToString().Trim()); //이관전작지라인
					oDS_PS_PP570L.SetValue("U_ColReg17", i, oRecordSet.Fields.Item("BPP030Seq").Value.ToString().Trim()); //이관전작업순서
					oDS_PS_PP570L.SetValue("U_ColReg18", i, oRecordSet.Fields.Item("AOrdNum").Value.ToString().Trim());	  //이관후작번
					oDS_PS_PP570L.SetValue("U_ColReg19", i, oRecordSet.Fields.Item("AOrdSub1").Value.ToString().Trim());  //이관후서브1
					oDS_PS_PP570L.SetValue("U_ColReg20", i, oRecordSet.Fields.Item("AOrdSub2").Value.ToString().Trim());  //이관후서브2
					oDS_PS_PP570L.SetValue("U_ColReg21", i, oRecordSet.Fields.Item("AItemName").Value.ToString().Trim()); //이관후품명
					oDS_PS_PP570L.SetValue("U_ColReg22", i, oRecordSet.Fields.Item("AItemSpec").Value.ToString().Trim()); //이관후규격

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
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
		/// PS_PP570_AddData 데이터 INSERT(UPDATE)
		/// </summary>
		private void PS_PP570_AddData()
		{
			int loopCount;
			string sQry;
			string errMessage = string.Empty;
			string CpCode;    //공정코드
			string BOrdNum;   //이관전작번
			string BOrdSub1;  //이관전서브작번1
			string BOrdSub2;  //이관전서브작번2
			string BPP030HNo; //이관전작지문서번호
			string BPP030MNo; //이관전작지라인
			string BPP030Seq; //이관전작업순서
			string AOrdNum;	  //이관후작번
			string AOrdSub1;  //이관후서브작번1
			string AOrdSub2;  //이관후서브작번2							  
			string PP040HNo;  //작업일보문서번호
			string PP040MNo;  //작업일보라인번호
			string TDocDate;  //이관일자
			string TCntcCode; //등록자사번
			string TCntcName; //등록자성명
			string RCntcCode; //요청자사번
			string RCntcName; //요청사성명
			string TransCls;  //이관구분

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				TDocDate = DateTime.Now.ToString("yyyyMMdd");
				TCntcCode = oForm.Items.Item("TCntcCode").Specific.Value.ToString().Trim();
				TCntcName = oForm.Items.Item("TCntcName").Specific.Value.ToString().Trim();
				RCntcCode = oForm.Items.Item("RCntcCode").Specific.Value.ToString().Trim();
				RCntcName = oForm.Items.Item("RCntcName").Specific.Value.ToString().Trim();
				TransCls = oForm.Items.Item("TransCls").Specific.Value.ToString().Trim();

				oMat.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					//체크 확인 로직 변경(FlushDataSource 후 에도 Check box 값을 못 가져옴, Matrix의 값을 직접 가져옴)
					if (oMat.Columns.Item("Check").Cells.Item(loopCount + 1).Specific.Checked == true)
					{
						CpCode    = oDS_PS_PP570L.GetValue("U_ColReg02", loopCount).ToString().Trim();
						BOrdNum   = oDS_PS_PP570L.GetValue("U_ColReg10", loopCount).ToString().Trim();
						BOrdSub1  = oDS_PS_PP570L.GetValue("U_ColReg11", loopCount).ToString().Trim();
						BOrdSub2  = oDS_PS_PP570L.GetValue("U_ColReg12", loopCount).ToString().Trim();
						BPP030HNo = oDS_PS_PP570L.GetValue("U_ColReg15", loopCount).ToString().Trim();
						BPP030MNo = oDS_PS_PP570L.GetValue("U_ColReg16", loopCount).ToString().Trim();
						BPP030Seq = oDS_PS_PP570L.GetValue("U_ColReg17", loopCount).ToString().Trim();
																				   
						AOrdNum   = oDS_PS_PP570L.GetValue("U_ColReg18", loopCount).ToString().Trim();
						AOrdSub1  = oDS_PS_PP570L.GetValue("U_ColReg19", loopCount).ToString().Trim();
						AOrdSub2  = oDS_PS_PP570L.GetValue("U_ColReg20", loopCount).ToString().Trim();
																				   
						PP040HNo  = oDS_PS_PP570L.GetValue("U_ColReg08", loopCount).ToString().Trim();
						PP040MNo  = oDS_PS_PP570L.GetValue("U_ColReg09", loopCount).ToString().Trim();

						if (TransCls == "01")
						{
							sQry = "EXEC [PS_PP570_02] '";
							sQry += CpCode + "','";
							sQry += BOrdNum + "','";
							sQry += BOrdSub1 + "','";
							sQry += BOrdSub2 + "','";
							sQry += BPP030HNo + "','";
							sQry += BPP030MNo + "','";
							sQry += BPP030Seq + "','";
							sQry += AOrdNum + "','";
							sQry += AOrdSub1 + "','";
							sQry += AOrdSub2 + "','";
							sQry += PP040HNo + "','";
							sQry += PP040MNo + "','";
							sQry += TDocDate + "','";
							sQry += TCntcCode + "','";
							sQry += TCntcName + "','";
							sQry += RCntcCode + "','";
							sQry += RCntcName + "'";
						}
						else
						{
							sQry = "EXEC [PS_PP570_03] '";
							sQry += CpCode + "','";
							sQry += BOrdNum + "','";
							sQry += BOrdSub1 + "','";
							sQry += BOrdSub2 + "','";
							sQry += BPP030HNo + "','";
							sQry += BPP030MNo + "','";
							sQry += BPP030Seq + "','";
							sQry += AOrdNum + "','";
							sQry += AOrdSub1 + "','";
							sQry += AOrdSub2 + "','";
							sQry += PP040HNo + "','";
							sQry += PP040MNo + "','";
							sQry += TDocDate + "','";
							sQry += TCntcCode + "','";
							sQry += TCntcName + "','";
							sQry += RCntcCode + "','";
							sQry += RCntcName + "'";
						}
						oRecordSet.DoQuery(sQry);

						//오류 발생 시 For문 종료
						if (oRecordSet.Fields.Item("ErrorMessage").Value != "True")
						{
							errMessage = loopCount + 1 + "행 " + oRecordSet.Fields.Item("ErrorMessage").Value;
							throw new Exception();
						}
					}
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_PP570_CheckAll
		/// </summary>
		private void PS_PP570_CheckAll()
		{
			string CheckType;
			short loopCount;

			try
			{
				oForm.Freeze(true);

				CheckType = "Y";
				oMat.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_PP570L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "N")
					{
						CheckType = "N";
						break;
					}
				}

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					oDS_PS_PP570L.Offset = loopCount;
					if (CheckType == "N")
					{
						oDS_PS_PP570L.SetValue("U_ColReg01", loopCount, "Y");
					}
					else
					{
						oDS_PS_PP570L.SetValue("U_ColReg01", loopCount, "N");
					}
				}
				oMat.LoadFromDataSource();
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
		/// PS_PP570_UpdateAfterOrdNum 이관후 작번을 매트릭스에 적용
		/// </summary>
		private void PS_PP570_UpdateAfterOrdNum()
		{
			short loopCount;
			string AOrdNum;	    //이관후 작번
			string AOrdSub1;	//이관후 서브1
			string AOrdSub2;	//이관후 서브2
			string AItemName;	//이관후 작명
			string AItemSpec;   //이관후 규격

			try
			{
				oForm.Freeze(true);

				AOrdNum   = oForm.Items.Item("AOrdNum").Specific.Value.ToString().Trim();
				AOrdSub1  = oForm.Items.Item("AOrdSub1").Specific.Value.ToString().Trim();
				AOrdSub2  = oForm.Items.Item("AOrdSub2").Specific.Value.ToString().Trim();
				AItemName = oForm.Items.Item("AItemName").Specific.Value.ToString().Trim();
				AItemSpec = oForm.Items.Item("AItemSpec").Specific.Value.ToString().Trim();

				oMat.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					//체크 확인 로직 변경(FlushDataSource 후 에도 Check box 값을 못 가져옴, Matrix의 값을 직접 가져옴)
					if (oMat.Columns.Item("Check").Cells.Item(loopCount + 1).Specific.Checked == true)
					{
						oDS_PS_PP570L.SetValue("U_ColReg18", loopCount, AOrdNum);
						oDS_PS_PP570L.SetValue("U_ColReg19", loopCount, AOrdSub1);
						oDS_PS_PP570L.SetValue("U_ColReg20", loopCount, AOrdSub2);
						oDS_PS_PP570L.SetValue("U_ColReg21", loopCount, AItemName);
						oDS_PS_PP570L.SetValue("U_ColReg22", loopCount, AItemSpec);
					}
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
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
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, pVal, BubbleEvent);
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
					if (pVal.ItemUID == "BtnSrch")
					{
						PS_PP570_MTX01();
					}
					else if (pVal.ItemUID == "BtnSave") //저장
					{
						if (PS_PP570_CheckDataValid() == false)
						{
							BubbleEvent = false;
							return;
						}
						else
						{
							if (PSH_Globals.SBO_Application.MessageBox("[" + oForm.Items.Item("BOrdNum").Specific.Value.ToString().Trim() + "] 작번의 비용이 [" + oForm.Items.Item("AOrdNum").Specific.Value.ToString().Trim() + "] 작번으로 이관됩니다. " + "             " + "이관후 취소는 불가능합니다. 정말로 저장하시겠습니까?", 1, "예", "아니오") == 1)
							{
								PS_PP570_AddData();
								PS_PP570_MTX01();
							}
							else
							{
								BubbleEvent = false;
								return;
							}
						}
					}
					else if (pVal.ItemUID == "BtnCfm")
					{
						PS_PP570_UpdateAfterOrdNum();
					}
					else if (pVal.ItemUID == "BtnAll") //전체선택(해제)
					{
						PS_PP570_CheckAll();
					}
					else if (pVal.BeforeAction == false)
					{
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "BOrdNum", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "AOrdNum", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "TCntcCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "RCntcCode", "");
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
							oMat.SelectRow(pVal.Row, true, false);
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
		}

		/// <summary>
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						PS_PP570_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP570L);
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
							PS_PP570_AddMatrixRow(oMat.VisualRowCount, false);
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
							oDS_PS_PP570L.RemoveRecord(oDS_PS_PP570L.Size - 1);
							oMat.LoadFromDataSource();
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
