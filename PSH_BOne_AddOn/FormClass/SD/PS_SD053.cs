using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 작번별미품의금액등록
	/// </summary>
	internal class PS_SD053 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;

		private SAPbouiCOM.DBDataSource oDS_PS_SD053L; //등록라인

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oLast_Mode;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD053.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD053_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD053");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_SD053_CreateItems();
				PS_SD053_ComboBox_Setting();
				PS_SD053_FormResize();
				PS_SD053_LoadCaption();
				PS_SD053_Initial_Setting();

				oForm.EnableMenu("1283", false);// 삭제
				oForm.EnableMenu("1286", false);// 닫기
				oForm.EnableMenu("1287", false);// 복제
				oForm.EnableMenu("1285", false);// 복원
				oForm.EnableMenu("1284", false);// 취소
				oForm.EnableMenu("1293", false);// 행삭제
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
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_SD053_CreateItems
		/// </summary>
		private void PS_SD053_CreateItems()
		{
			try
			{
				oDS_PS_SD053L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//입력정보
				//작번
				oForm.DataSources.UserDataSources.Add("OrdNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("OrdNum").Specific.DataBind.SetBound(true, "", "OrdNum");

				//품명
				oForm.DataSources.UserDataSources.Add("FrgnName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("FrgnName").Specific.DataBind.SetBound(true, "", "FrgnName");

				//기준년월
				oForm.DataSources.UserDataSources.Add("StdYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("StdYM").Specific.DataBind.SetBound(true, "", "StdYM");

				//기준회차
				oForm.DataSources.UserDataSources.Add("StdCnt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("StdCnt").Specific.DataBind.SetBound(true, "", "StdCnt");

				//조회정보
				//작번
				oForm.DataSources.UserDataSources.Add("OrdNumS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("OrdNumS").Specific.DataBind.SetBound(true, "", "OrdNumS");

				//품명
				oForm.DataSources.UserDataSources.Add("FrgnNameS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("FrgnNameS").Specific.DataBind.SetBound(true, "", "FrgnNameS");

				//기준년월
				oForm.DataSources.UserDataSources.Add("StdYMS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("StdYMS").Specific.DataBind.SetBound(true, "", "StdYMS");

				//기준회차
				oForm.DataSources.UserDataSources.Add("StdCntS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("StdCntS").Specific.DataBind.SetBound(true, "", "StdCntS");

				//계
				oForm.DataSources.UserDataSources.Add("Total", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("Total").Specific.DataBind.SetBound(true, "", "Total");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD053_ComboBox_Setting
		/// </summary>
		private void PS_SD053_ComboBox_Setting()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//입력정보
				//기준회차
				oForm.Items.Item("StdCnt").Specific.ValidValues.Add("%", "선택");
				sQry = "  SELECT      U_Minor AS [Code],";
				sQry += "                U_CdName AS [Name]";
				sQry += " FROM       [@PS_SY001L]";
				sQry += " WHERE      Code = 'S008'";
				sQry += "                AND U_UseYN = 'Y'";
				sQry += " ORDER BY  U_Seq";
				dataHelpClass.Set_ComboList(oForm.Items.Item("StdCnt").Specific, sQry, "", false, false);
				oForm.Items.Item("StdCnt").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//조회정보
				//기준회차
				oForm.Items.Item("StdCntS").Specific.ValidValues.Add("%", "전체");
				sQry = "  SELECT      U_Minor AS [Code],";
				sQry += "                U_CdName AS [Name]";
				sQry += " FROM       [@PS_SY001L]";
				sQry += " WHERE      Code = 'S008'";
				sQry += "                AND U_UseYN = 'Y'";
				sQry += " ORDER BY  U_Seq";
				dataHelpClass.Set_ComboList(oForm.Items.Item("StdCntS").Specific, sQry, "", false, false);
				oForm.Items.Item("StdCntS").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//매트릭스
				//기준회차
				sQry = "    SELECT      U_Minor AS [Code],";
				sQry += "                U_CdName AS [Name]";
				sQry += " FROM       [@PS_SY001L]";
				sQry += " WHERE      Code = 'S008'";
				sQry += "                AND U_UseYN = 'Y'";
				sQry += " ORDER BY  U_Seq";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("StdCnt"), sQry, "", "");

				//품의구분
				sQry = "    SELECT       Code AS [Code],";
				sQry += "                 Name AS [Name]";
				sQry += " FROM        [@PSH_ORDTYP]";
				sQry += " WHERE       Code IN ('10','20','30','40')";   //4개 품의대해서만 조회
				sQry += " ORDER BY   Code";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("POType"), sQry, "", "");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD053_FormResize
		/// </summary>
		private void PS_SD053_FormResize()
		{
			try
			{
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD053_LoadCaption Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
		/// </summary>
		private void PS_SD053_LoadCaption()
		{
			try
			{
				oForm.Freeze(true);

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
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
		/// PS_SD053_Initial_Setting
		/// </summary>
		private void PS_SD053_Initial_Setting()
		{
			try
			{
				oMat.Columns.Item("StdYM").Visible = false;
				oMat.Columns.Item("StdCnt").Visible = false;
				oForm.Items.Item("StdYM").Specific.Value = DateTime.Now.ToString("yyyyMM");
				oForm.Items.Item("StdYMS").Specific.Value = DateTime.Now.ToString("yyyyMM");
				oForm.Items.Item("OrdNum").Click();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD053_CheckAll
		/// </summary>
		private void PS_SD053_CheckAll()
		{
			string CheckType;
			int loopCount;

			try
			{
				oForm.Freeze(true);

				CheckType = "Y";
				oMat.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_SD053L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "N")
					{
						CheckType = "N";
						break; // TODO: might not be correct. Was : Exit For
					}
				}

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					oDS_PS_SD053L.Offset = loopCount;
					if (CheckType == "N")
					{
						oDS_PS_SD053L.SetValue("U_ColReg01", loopCount, "Y");
					}
					else
					{
						oDS_PS_SD053L.SetValue("U_ColReg01", loopCount, "N");
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
		/// PS_SD053_CheckBeforeSearch 필수입력사항 체크
		/// </summary>
		/// <param name="pItemUID"></param>
		/// <returns></returns>
		private bool PS_SD053_CheckBeforeSearch(string pItemUID)
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (pItemUID == "BtnSearch1")
				{
					if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim()))
					{
						errMessage = "입력정보의 작번은 필수사항입니다. 확인하세요.";
						throw new Exception();
					}
					if (string.IsNullOrEmpty(oForm.Items.Item("StdYM").Specific.Value.ToString().Trim()))
					{
						errMessage = "입력정보의 기준년월은 필수사항입니다. 확인하세요.";
						throw new Exception();
					}
					if (oForm.Items.Item("StdCnt").Specific.Value.ToString().Trim() == "%")
					{
						errMessage = "입력정보의 기준회차는 필수사항입니다. 확인하세요.";
						throw new Exception();
					}
				}
				else if (pItemUID == "BtnSearch2")
				{
					if (string.IsNullOrEmpty(oForm.Items.Item("OrdNumS").Specific.Value.ToString().Trim()))
					{
						errMessage = "조회정보의 작번은 필수사항입니다. 확인하세요.";
						throw new Exception();
					}
					if (string.IsNullOrEmpty(oForm.Items.Item("StdYMS").Specific.Value.ToString().Trim()))
					{
						errMessage = "조회정보의 기준년월은 필수사항입니다. 확인하세요.";
						throw new Exception();
					}
					if (oForm.Items.Item("StdCntS").Specific.Value.ToString().Trim() == "%")
					{
						errMessage = "조회정보의 기준회차는 필수사항입니다. 확인하세요.";
						throw new Exception();
					}
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
		/// PS_SD053_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_SD053_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			int loopCount;
			double TotalAmt = 0;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "Mat01":
						if (oCol == "Amount")
						{
							oMat.FlushToDataSource();

							for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
							{
								TotalAmt += Convert.ToDouble(oDS_PS_SD053L.GetValue("U_ColSum03", loopCount).ToString().Trim());
							}
							oForm.Items.Item("Total").Specific.Value = TotalAmt;
							oMat.LoadFromDataSource();
						}
						oMat.AutoResizeColumns();
						break;

					case "OrdNum":
						oForm.Items.Item("FrgnName").Specific.Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
						break;

					case "OrdNumS":
						oForm.Items.Item("FrgnNameS").Specific.Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD053_MTX01 데이터 조회
		/// </summary>
		/// <param name="pItemUID"></param>
		private void PS_SD053_MTX01(string pItemUID)
		{
			int i;
			string sQry;
			string errMessage = string.Empty;

			string OrdNum;       //작번
			string StdYM;        //기준년월
			string StdCnt;       //기준회차
			string CntcCode;     //사용자 사번
			double TotalAmt = 0; //금액 합계

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				ProgressBar01.Text = "조회시작!";

				if (pItemUID == "BtnSearch1")
				{
					OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
					StdYM = oForm.Items.Item("StdYM").Specific.Value.ToString().Trim();
					StdCnt = oForm.Items.Item("StdCnt").Specific.Value.ToString().Trim();
					CntcCode = dataHelpClass.User_MSTCOD();

					sQry = " EXEC [PS_SD053_01] '";
					sQry += OrdNum + "','";
					sQry += StdYM + "','";
					sQry += StdCnt + "','";
					sQry += CntcCode + "'";
					oRecordSet.DoQuery(sQry);

					oMat.Clear();
					oDS_PS_SD053L.Clear();
					oMat.FlushToDataSource();
					oMat.LoadFromDataSource();

					if (oRecordSet.RecordCount == 0)
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_SD053_LoadCaption();
						errMessage = "결과가 존재하지 않습니다.";
						throw new Exception();
					}

					for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
					{
						if (i + 1 > oDS_PS_SD053L.Size)
						{
							oDS_PS_SD053L.InsertRecord(i);
						}

						oMat.AddRow();
						oDS_PS_SD053L.Offset = i;

						oDS_PS_SD053L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
						oDS_PS_SD053L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Check").Value.ToString().Trim());         //선택
						oDS_PS_SD053L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("StdYM").Value.ToString().Trim());         //기준년월
						oDS_PS_SD053L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("StdCnt").Value.ToString().Trim());        //기준회차
						oDS_PS_SD053L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("MM005DocEntry").Value.ToString().Trim()); //요청번호
						oDS_PS_SD053L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("MainOrdNum").Value.ToString().Trim());    //품목코드(작번)
						oDS_PS_SD053L.SetValue("U_ColReg18", i, oRecordSet.Fields.Item("SubNo1").Value.ToString().Trim());        //서브작번1
						oDS_PS_SD053L.SetValue("U_ColReg19", i, oRecordSet.Fields.Item("SubNo2").Value.ToString().Trim());        //서브작번2
						oDS_PS_SD053L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("MainFrgnName").Value.ToString().Trim());  //품목명(작번)
						oDS_PS_SD053L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("POTypeCD").Value.ToString().Trim());      //품의구분
						oDS_PS_SD053L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("SItemCode").Value.ToString().Trim());     //품목코드(자재)
						oDS_PS_SD053L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("SItemName").Value.ToString().Trim());     //품목명(자재)
						oDS_PS_SD053L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("SItemSpec").Value.ToString().Trim());     //규격(자재)
						oDS_PS_SD053L.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("MM005Weight").Value.ToString().Trim());   //요청수량
						oDS_PS_SD053L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("ResPrice").Value.ToString().Trim());      //실적단가
						oDS_PS_SD053L.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("ResAmt").Value.ToString().Trim());        //실적금액
						oDS_PS_SD053L.SetValue("U_ColSum04", i, oRecordSet.Fields.Item("PreAmount").Value.ToString().Trim());     //직전예상금액
						oDS_PS_SD053L.SetValue("U_ColSum03", i, oRecordSet.Fields.Item("Amount").Value.ToString().Trim());        //예상금액
						oDS_PS_SD053L.SetValue("U_ColTxt01", i, oRecordSet.Fields.Item("Comment").Value.ToString().Trim());       //비고
						oDS_PS_SD053L.SetValue("U_ColReg15", i, oRecordSet.Fields.Item("CreateUser").Value.ToString().Trim());    //등록자(사번)
						oDS_PS_SD053L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item("CreateDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //등록일자
						oDS_PS_SD053L.SetValue("U_ColReg17", i, oRecordSet.Fields.Item("UpdateUser").Value.ToString().Trim());    //수정자(사번)	  
						oDS_PS_SD053L.SetValue("U_ColDt02", i, Convert.ToDateTime(oRecordSet.Fields.Item("UpdateDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //수정일자

						TotalAmt += Convert.ToDouble(oRecordSet.Fields.Item("Amount").Value.ToString().Trim());

						oRecordSet.MoveNext();
						ProgressBar01.Value += 1;
						ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
					}
					oForm.Items.Item("Total").Specific.Value = TotalAmt;
					oMat.LoadFromDataSource();
					oMat.AutoResizeColumns();
				}
				else if (pItemUID == "BtnSearch2")
				{
					OrdNum = oForm.Items.Item("OrdNumS").Specific.Value.ToString().Trim();
					StdYM = oForm.Items.Item("StdYMS").Specific.Value.ToString().Trim();
					StdCnt = oForm.Items.Item("StdCntS").Specific.Value.ToString().Trim();
					CntcCode = dataHelpClass.User_MSTCOD();

					sQry = "  EXEC [PS_SD053_02] '";
					sQry += OrdNum + "','";
					sQry += StdYM + "','";
					sQry += StdCnt + "','";
					sQry += CntcCode + "'";
					oRecordSet.DoQuery(sQry);

					oMat.Clear();
					oDS_PS_SD053L.Clear();
					oMat.FlushToDataSource();
					oMat.LoadFromDataSource();

					if (oRecordSet.RecordCount == 0)
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_SD053_LoadCaption();
						errMessage = "결과가 존재하지 않습니다.";
						throw new Exception();
					}

					for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
					{
						if (i + 1 > oDS_PS_SD053L.Size)
						{
							oDS_PS_SD053L.InsertRecord(i);
						}

						oMat.AddRow();
						oDS_PS_SD053L.Offset = i;

						oDS_PS_SD053L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
						oDS_PS_SD053L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Check").Value.ToString().Trim());       //선택
						oDS_PS_SD053L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("StdYM").Value.ToString().Trim());       //기준년월
						oDS_PS_SD053L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("StdCnt").Value.ToString().Trim());      //기준회차
						oDS_PS_SD053L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("ReqNo").Value.ToString().Trim());       //요청번호
						oDS_PS_SD053L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());      //품목코드(작번)
						oDS_PS_SD053L.SetValue("U_ColReg18", i, oRecordSet.Fields.Item("SubNo1").Value.ToString().Trim());      //서브작번1
						oDS_PS_SD053L.SetValue("U_ColReg19", i, oRecordSet.Fields.Item("SubNo2").Value.ToString().Trim());      //서브작번2
						oDS_PS_SD053L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("FrgnName").Value.ToString().Trim());    //품목명(작번)
						oDS_PS_SD053L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("POType").Value.ToString().Trim());      //품의구분
						oDS_PS_SD053L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());    //품목코드(자재)
						oDS_PS_SD053L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());    //품목명(자재)
						oDS_PS_SD053L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim());    //규격(자재)
						oDS_PS_SD053L.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("ReqQty").Value.ToString().Trim());      //요청수량
						oDS_PS_SD053L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("ResPrice").Value.ToString().Trim());    //실적단가
						oDS_PS_SD053L.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("ResAmt").Value.ToString().Trim());      //실적금액
						oDS_PS_SD053L.SetValue("U_ColSum04", i, oRecordSet.Fields.Item("PreAmount").Value.ToString().Trim());   //직전예상금액
						oDS_PS_SD053L.SetValue("U_ColSum03", i, oRecordSet.Fields.Item("Amount").Value.ToString().Trim());      //예상금액
						oDS_PS_SD053L.SetValue("U_ColTxt01", i, oRecordSet.Fields.Item("Comment").Value.ToString().Trim());     //비고
						oDS_PS_SD053L.SetValue("U_ColReg15", i, oRecordSet.Fields.Item("CreateUser").Value.ToString().Trim());  //등록자(사번)
						oDS_PS_SD053L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item("CreateDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //등록일자
						oDS_PS_SD053L.SetValue("U_ColReg17", i, oRecordSet.Fields.Item("UpdateUser").Value.ToString().Trim());  //수정자(사번)
						oDS_PS_SD053L.SetValue("U_ColDt02", i, Convert.ToDateTime(oRecordSet.Fields.Item("UpdateDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //수정일자

						TotalAmt += Convert.ToDouble(oRecordSet.Fields.Item("Amount").Value.ToString().Trim());

						oRecordSet.MoveNext();
						ProgressBar01.Value += 1;
						ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
					}
					oForm.Items.Item("Total").Specific.Value = TotalAmt;
					oMat.LoadFromDataSource();
					oMat.AutoResizeColumns();
				}
			}
			catch (Exception ex)
			{
				ProgressBar01.Stop();  //stop 안하면 오래결림.

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
		/// PS_SD053_AddData  데이터 INSERT
		/// </summary>
		/// <returns></returns>
		private bool PS_SD053_AddData()
		{
			bool functionReturnValue = false;

			int loopCount;
			string sQry;
							 
			string StdYM;	  //기준년월
			string StdCnt;	  //기준회차
			string ReqNo;	  //요청번호
			string OrdNum;	  //품목코드(작번)
			string SubNo1;	  //서브작번1
			string SubNo2;	  //서브작번2
			string FrgnName;  //품목명(작번)
			string POType;	  //품의구분
			string ItemCode;  //품목코드(자재)
			string ItemName;  //품목명(자재)
			string ItemSpec;  //규격(자재)
			double ReqQty;	  //요청수량
			double ResPrice;  //실적단가
			double ResAmt;	  //실적금액
			double PreAmount; //직전예상금액
			double Amount;	  //예상금액
			string Comment;	  //비고
			string CntcCode;  //등록자 및 수정자

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				CntcCode = dataHelpClass.User_MSTCOD();	//사용자사번

				oMat.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_SD053L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						StdYM  = oDS_PS_SD053L.GetValue("U_ColReg02", loopCount).ToString().Trim();
						StdCnt = oDS_PS_SD053L.GetValue("U_ColReg03", loopCount).ToString().Trim();
						ReqNo  = oDS_PS_SD053L.GetValue("U_ColReg04", loopCount).ToString().Trim();
						OrdNum = oDS_PS_SD053L.GetValue("U_ColReg05", loopCount).ToString().Trim();
						SubNo1 = oDS_PS_SD053L.GetValue("U_ColReg18", loopCount).ToString().Trim();
						SubNo2 = oDS_PS_SD053L.GetValue("U_ColReg19", loopCount).ToString().Trim();
						FrgnName = oDS_PS_SD053L.GetValue("U_ColReg06", loopCount).ToString().Trim();
						POType = oDS_PS_SD053L.GetValue("U_ColReg07", loopCount).ToString().Trim();
						ItemCode = oDS_PS_SD053L.GetValue("U_ColReg08", loopCount).ToString().Trim();
						ItemName = oDS_PS_SD053L.GetValue("U_ColReg09", loopCount).ToString().Trim();
						ItemSpec = oDS_PS_SD053L.GetValue("U_ColReg10", loopCount).ToString().Trim();
						ReqQty = Convert.ToDouble(oDS_PS_SD053L.GetValue("U_ColQty01", loopCount).ToString().Trim());
						ResPrice = Convert.ToDouble(oDS_PS_SD053L.GetValue("U_ColSum01", loopCount).ToString().Trim());
						ResAmt = Convert.ToDouble(oDS_PS_SD053L.GetValue("U_ColSum02", loopCount).ToString().Trim());
						PreAmount = Convert.ToDouble(oDS_PS_SD053L.GetValue("U_ColSum04", loopCount).ToString().Trim());
						Amount = Convert.ToDouble(oDS_PS_SD053L.GetValue("U_ColSum03", loopCount).ToString().Trim());
						Comment = oDS_PS_SD053L.GetValue("U_ColTxt01", loopCount).ToString().Trim();

						sQry = " EXEC [PS_SD053_03] ";
						sQry += "'" + StdYM + "',";
						sQry += "'" + StdCnt + "',";
						sQry += "'" + ReqNo + "',";
						sQry += "'" + OrdNum + "',";
						sQry += "'" + SubNo1 + "',";
						sQry += "'" + SubNo2 + "',";
						sQry += "'" + FrgnName + "',";
						sQry += "'" + POType + "',";
						sQry += "'" + ItemCode + "',";
						sQry += "'" + ItemName + "',";
						sQry += "'" + ItemSpec + "',";
						sQry += "'" + ReqQty + "',";
						sQry += "'" + ResPrice + "',";
						sQry += "'" + ResAmt + "',";
						sQry += "'" + PreAmount + "',";
						sQry += "'" + Amount + "',";
						sQry += "'" + Comment + "',";
						sQry += "'" + CntcCode + "'";
						oRecordSet.DoQuery(sQry);
					}
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				functionReturnValue = true;
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
			return functionReturnValue;
		}

		/// <summary>
		/// PS_SD053_DeleteData 기본정보 삭제
		/// </summary>
		private void PS_SD053_DeleteData()
		{
			int loopCount;
			string sQry;
			string errMessage = string.Empty;

			string StdYM;	//기준년월
			string StdCnt;	//기준회차
			string ReqNo;   //요청번호

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (oMat.VisualRowCount == 0)
				{
					errMessage = "삭제대상이 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{

					StdYM = oDS_PS_SD053L.GetValue("U_ColReg02", loopCount).ToString().Trim();
					StdCnt = oDS_PS_SD053L.GetValue("U_ColReg03", loopCount).ToString().Trim();
					ReqNo = oDS_PS_SD053L.GetValue("U_ColReg04", loopCount).ToString().Trim();

					sQry = " EXEC [PS_SD053_04] ";
					sQry += "'" + StdYM + "',";
					sQry += "'" + StdCnt + "',";
					sQry += "'" + ReqNo + "'";
					oRecordSet.DoQuery(sQry);
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("삭제 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
                    Raise_EVENT_FORM_RESIZE(FormUID, pVal, BubbleEvent);
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
					if (pVal.ItemUID == "BtnAdd")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_SD053_AddData() == false)
							{
								BubbleEvent = false;
								return;
							}
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_SD053_LoadCaption();
							oLast_Mode = Convert.ToInt32(oForm.Mode);
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_SD053_AddData() == false)
							{
								BubbleEvent = false;
								return;
							}
							PS_SD053_MTX01("BtnSearch2");
							PS_SD053_LoadCaption();
						}
					}
					else if (pVal.ItemUID == "BtnSearch1")
					{
						if (PS_SD053_CheckBeforeSearch(pVal.ItemUID) == false)
						{
							BubbleEvent = false;
							return;
						}
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_SD053_LoadCaption();
						PS_SD053_MTX01(pVal.ItemUID);
					}
					else if (pVal.ItemUID == "BtnSearch2")
					{
						if (PS_SD053_CheckBeforeSearch(pVal.ItemUID) == false)
						{
							BubbleEvent = false;
							return;
						}
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
						PS_SD053_LoadCaption();
						PS_SD053_MTX01(pVal.ItemUID);
					}
					else if (pVal.ItemUID == "BtnDelete")
					{

						if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
						{
							PS_SD053_DeleteData();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_SD053_LoadCaption();
						}
					}
					else if (pVal.ItemUID == "BtnChk")
					{
						PS_SD053_CheckAll();
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "BtnAdd" | pVal.ItemUID == "BtnDelete")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							oForm.Items.Item("Total").Specific.Value = 0;

						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
						}
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OrdNum", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OrdNumS", "");
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
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oMat.SelectRow(pVal.Row, true, false);
						}
					}
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
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "Mat01")
						{
							PS_SD053_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
						else
						{
							PS_SD053_FlushToItemValue(pVal.ItemUID, 0, "");
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
		/// Raise_EVENT_FORM_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_RESIZE(string FormUID, SAPbouiCOM.ItemEvent pVal, bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_SD053_FormResize();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD053L);
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
							//추가버튼 클릭시 메트릭스 insertrow
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							BubbleEvent = false;
							PS_SD053_LoadCaption();
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "7169": //엑셀 내보내기
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

		/// <summary>
		/// FormDataEvent
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
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
							break;
					}
				}
				else if (BusinessObjectInfo.BeforeAction == false)
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
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
