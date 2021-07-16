using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 장비별 공정진행률 등록
	/// </summary>
	internal class PS_PP568 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP568A; //등록용 Matrix

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP568.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP568_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP568");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP568_CreateItems();
				PS_PP568_ComboBox_Setting();
				PS_PP568_Initial_Setting();
				PS_PP568_LoadCaption();

				oForm.EnableMenu(("1283"), false); // 삭제
				oForm.EnableMenu(("1286"), false); // 닫기
				oForm.EnableMenu(("1287"), false); // 복제
				oForm.EnableMenu(("1285"), false); // 복원
				oForm.EnableMenu(("1284"), false); // 취소
				oForm.EnableMenu(("1293"), false); // 행삭제
				oForm.EnableMenu(("1281"), false);
				oForm.EnableMenu(("1282"), true);
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
		/// PS_PP568_CreateItems
		/// </summary>
		private void PS_PP568_CreateItems()
		{
			try
			{
				oDS_PS_PP568A = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//수주년월(시작)
				oForm.DataSources.UserDataSources.Add("FrYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("FrYM").Specific.DataBind.SetBound(true, "", "FrYM");

				//수주년월(종료)
				oForm.DataSources.UserDataSources.Add("ToYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("ToYM").Specific.DataBind.SetBound(true, "", "ToYM");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");

				//수주처
				oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");

				//수주처명
				oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");

				//작번
				oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

				//품명
				oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

				//규격
				oForm.DataSources.UserDataSources.Add("ItemSpec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemSpec").Specific.DataBind.SetBound(true, "", "ItemSpec");

				//자체/외주
				oForm.DataSources.UserDataSources.Add("InOut", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("InOut").Specific.DataBind.SetBound(true, "", "InOut");

				//수주금액
				oForm.DataSources.UserDataSources.Add("OrdAmt", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("OrdAmt").Specific.DataBind.SetBound(true, "", "OrdAmt");

				//생산완료포함CheckBox
				oForm.DataSources.UserDataSources.Add("CmltYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("CmltYN").Specific.DataBind.SetBound(true, "", "CmltYN");

				//등록일
				oForm.DataSources.UserDataSources.Add("RegDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("RegDt").Specific.DataBind.SetBound(true, "", "RegDt");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP568_ComboBox_Setting
		/// </summary>
		private void PS_PP568_ComboBox_Setting()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				//거래처구분
				sQry = " SELECT      U_Minor,";
				sQry += "             U_CdName";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'C100'";
				sQry += "             AND U_UseYN = 'Y'";
				sQry += " ORDER BY    U_Seq";
				oForm.Items.Item("CardType").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, sQry, "", false, false);
				oForm.Items.Item("CardType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//자체/외주
				oForm.Items.Item("InOut").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("InOut").Specific.ValidValues.Add("IN", "자체");
				oForm.Items.Item("InOut").Specific.ValidValues.Add("OUT", "외주");
				oForm.Items.Item("InOut").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP568_Initial_Setting
		/// </summary>
		private void PS_PP568_Initial_Setting()
		{
			try
			{
				oForm.Items.Item("FrYM").Specific.VALUE = DateTime.Now.AddMonths(-6).ToString("yyyyMM"); //6개월전
				oForm.Items.Item("ToYM").Specific.VALUE = DateTime.Now.ToString("yyyyMM");
				oForm.Items.Item("OrdAmt").Specific.VALUE = "100000000";
				oForm.Items.Item("RegDt").Specific.VALUE = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("CardCode").Click();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP568_LoadCaption Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
		/// </summary>
		private void PS_PP568_LoadCaption()
		{
			try
			{
				oForm.Freeze(true);

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BtnAdd").Specific.Caption = "저장";
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
		/// PS_PP568_Add_MatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP568_Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP568A.InsertRecord((oRow));
				}
				oMat.AddRow();
				oDS_PS_PP568A.Offset = oRow;
				oDS_PS_PP568A.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oDS_PS_PP568A.SetValue("U_ColReg01", oRow, "Y");

				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP568_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP568_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "CardCode": //거래처
						oForm.Items.Item("CardName").Specific.VALUE = dataHelpClass.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'", "");
						break;

					case "ItemCode": //작번,규격
						oForm.Items.Item("ItemName").Specific.VALUE = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'", "");
						oForm.Items.Item("ItemSpec").Specific.VALUE = dataHelpClass.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'", ""); 
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP568_FormReset 화면 초기화
		/// </summary>
		private void PS_PP568_FormReset()
		{
			string User_BPLId;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				
				User_BPLId = dataHelpClass.User_BPLID();
				oForm.DataSources.UserDataSources.Item("StdDate").Value = DateTime.Now.ToString("yyyyMMdd"); //기준일자

				oMat.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();
				PS_PP568_Add_MatrixRow(0, true);
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
		/// PS_PP568_CheckAll
		/// </summary>
		private void PS_PP568_CheckAll()
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
					if (oDS_PS_PP568A.GetValue("U_ColReg01", loopCount).ToString().Trim() == "N")
					{
						CheckType = "N";
						break;
					}
				}

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					oDS_PS_PP568A.Offset = loopCount;
					if (CheckType == "N")
					{
						oDS_PS_PP568A.SetValue("U_ColReg01", loopCount, "Y");
					}
					else
					{
						oDS_PS_PP568A.SetValue("U_ColReg01", loopCount, "N");
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
		/// PS_PP568_MTX01 등록된 데이터 조회
		/// </summary>
		private void PS_PP568_MTX01()
		{
			short i;
			string sQry;
			string errMessage = string.Empty;

			string FrYM;	 //기준년월(시작)
			string ToYM;	 //기준년월(종료)
			string CardType; //거래처구분
			string CardCode; //수주처
			string ItemCode; //작번
			string InOut;	 //자체/외주
			double OrdAmt;	 //수주금액
			string CmltYN;   //생산완료

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				FrYM = oForm.Items.Item("FrYM").Specific.Value.ToString().Trim();
				ToYM     = oForm.Items.Item("ToYM").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType").Specific.Selected.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				InOut    = oForm.Items.Item("InOut").Specific.Selected.Value.ToString().Trim();
				OrdAmt = Convert.ToDouble(oForm.Items.Item("OrdAmt").Specific.Value.ToString().Trim());

				if (oForm.Items.Item("CmltYN").Specific.Checked == true)
				{
					CmltYN = "Y";
				}
				else
				{
					CmltYN = "N";
				}

				ProgressBar01.Text = "조회중...";

				sQry = " EXEC [PS_PP568_01] '";
				sQry += FrYM + "','";
				sQry += ToYM + "','";
				sQry += CardType + "','";
				sQry += CardCode + "','";
				sQry += ItemCode + "','";
				sQry += InOut + "','";
				sQry += OrdAmt + "','";
				sQry += CmltYN + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_PP568A.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if ((oRecordSet.RecordCount == 0))
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_PP568_Add_MatrixRow(0, true);
					PS_PP568_LoadCaption();

					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP568A.Size)
					{
						oDS_PS_PP568A.InsertRecord((i));
					}

					oMat.AddRow();
					oDS_PS_PP568A.Offset = i;

					oDS_PS_PP568A.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP568A.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Select").Value.ToString().Trim());	 //선택
					oDS_PS_PP568A.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim()); //작번
					oDS_PS_PP568A.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim()); //작명
					oDS_PS_PP568A.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim()); //규격
					oDS_PS_PP568A.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("PreRegDt").Value.ToString().Trim()); //직전등록일
					oDS_PS_PP568A.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("PreRate").Value.ToString().Trim());	 //직전공정률
					oDS_PS_PP568A.SetValue("U_ColQty02", i, oRecordSet.Fields.Item("CurRate").Value.ToString().Trim());	 //현재공정률

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
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP568_DeleteData 기본정보 삭제
		/// </summary>
		private void PS_PP568_DeleteData()
		{
			short loopCount;
			string sQry;
			
			string ItemCode;
			string RegDt;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oMat.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_PP568A.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						ItemCode = oDS_PS_PP568A.GetValue("U_ColReg02", loopCount).ToString().Trim();
						RegDt = Convert.ToDateTime(oDS_PS_PP568A.GetValue("U_ColReg05", loopCount).ToString().Trim()).ToString("yyyyMMdd");

						sQry = " EXEC [PS_PP568_03] '";
						sQry += ItemCode + "','";     //작번
						sQry += RegDt + "'";          //등록일
							 
						oRecordSet.DoQuery(sQry);
					}
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("삭제완료", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
		/// PS_PP568_AddData 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
		/// </summary>
		/// <returns></returns>
		private bool PS_PP568_AddData()
		{
			bool functionReturnValue = false;

			short i;
			string sQry;
			string errMessage = string.Empty;

			string ItemCode; //작번
			string RegDt;	 //등록일
			double CpRate;   //공정진도율

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				RegDt = oForm.DataSources.UserDataSources.Item("RegDt").Value.ToString().Trim();

				oMat.FlushToDataSource();   //마지막 빈행 제외를 위해 2를 뺌

				ProgressBar01.Text = "저장중...";

				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					if (oDS_PS_PP568A.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
					{
						ItemCode = oDS_PS_PP568A.GetValue("U_ColReg02", i).ToString().Trim();				
						CpRate = Convert.ToDouble(oDS_PS_PP568A.GetValue("U_ColQty02", i).ToString().Trim());

						sQry = " EXEC [PS_PP568_02] '";
						sQry += ItemCode + "','";
						sQry += RegDt + "','";
						sQry += CpRate + "'";
						oRecordSet.DoQuery(sQry);

						ProgressBar01.Value += 1;
						ProgressBar01.Text = ProgressBar01.Value + "/" + Convert.ToString(oMat.VisualRowCount - 1) + "건 저장중...";
					}
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("저장완료", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
			return functionReturnValue;
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
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "BtnAdd") //추가/확인 버튼클릭
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_PP568_AddData() == false)
							{
								BubbleEvent = false;
								return;
							}
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

							PS_PP568_LoadCaption();
							PS_PP568_MTX01();

							oLast_Mode = Convert.ToInt32(oForm.Mode);

						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							PS_PP568_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

							PS_PP568_LoadCaption();
							PS_PP568_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnSearch") //조회
					{
						PS_PP568_MTX01();
					}
					else if (pVal.ItemUID == "BtnDelete") //삭제
					{
						if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
						{
							PS_PP568_DeleteData();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_PP568_LoadCaption();
							PS_PP568_MTX01();
						}
						else
						{
						}
					}
					else if (pVal.ItemUID == "BtnSelect") //전체선택
					{
						PS_PP568_CheckAll();
					}
					else if (pVal.ItemUID == "BtnPrt01")
					{
						// 출력없음
						//System.Threading.Thread thread = new System.Threading.Thread(PS_PP568_Print_Report01);
						//thread.SetApartmentState(System.Threading.ApartmentState.STA);
						//thread.Start();
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", "");
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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
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
						PS_PP568_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
					}
					else
					{
						PS_PP568_FlushToItemValue(pVal.ItemUID, 0, "");
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
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						if ((pVal.ItemUID == "Mat01"))
						{
							PS_PP568_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
						else
						{
							PS_PP568_FlushToItemValue(pVal.ItemUID, 0, "");
						}
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
					oMat.AutoResizeColumns();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP568A);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_ROW_DELETE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
		{
			int i;

			try
			{
				if ((oLastColRow01 > 0))
				{
					if (pVal.BeforeAction == true)
					{
					}
					else if (pVal.BeforeAction == false)
					{
						for (i = 1; i <= oMat.VisualRowCount; i++)
						{
							oMat.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
						}
						oMat.FlushToDataSource();
						oDS_PS_PP568A.RemoveRecord(oDS_PS_PP568A.Size - 1);
						oMat.LoadFromDataSource();

						if (oMat.RowCount == 0)
						{
							PS_PP568_Add_MatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_PP568A.GetValue("U_ColReg01", oMat.RowCount - 1).ToString().Trim()))
							{
								PS_PP568_Add_MatrixRow(oMat.RowCount, false);
							}
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
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_PP568_FormReset();
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
							PS_PP568_Add_MatrixRow(oMat.VisualRowCount, false);
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
							oDS_PS_PP568A.RemoveRecord(oDS_PS_PP568A.Size - 1);
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
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:    //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:     //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:  //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:  //36
							break;
					}
				}
				else if (BusinessObjectInfo.BeforeAction == false)
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:    //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:     //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:  //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:  //36
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
