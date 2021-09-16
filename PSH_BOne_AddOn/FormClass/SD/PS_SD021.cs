using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 판매계획등록
	/// </summary>
	internal class PS_SD021 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_SD021H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD021L; //등록라인
		
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		private string oDocEntry01;
		private SAPbouiCOM.BoFormMode oFormMode01;

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD021.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD021_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD021");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry"; 

				oForm.Freeze(true);

				PS_SD021_CreateItems();
				PS_SD021_ComboBox_Setting();
				PS_SD021_EnableMenus();
				PS_SD021_SetDocument(oFormDocEntry);

				oForm.Items.Item("StdYear").Click();
				
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
		/// PS_SD021_CreateItems
		/// </summary>
		private void PS_SD021_CreateItems()
		{
			try
			{
				oDS_PS_SD021H = oForm.DataSources.DBDataSources.Item("@PS_SD021H");
				oDS_PS_SD021L = oForm.DataSources.DBDataSources.Item("@PS_SD021L");
				oMat = oForm.Items.Item("Mat01").Specific;

				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oMat.AutoResizeColumns();

				oForm.Items.Item("StdYear").Specific.Value = DateTime.Now.ToString("yyyy");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD021_ComboBox_Setting
		/// </summary>
		private void PS_SD021_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD021_EnableMenus 메뉴활성화
		/// </summary>
		private void PS_SD021_EnableMenus()
		{
			try
			{
				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", true);  // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", true);  // 행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD021_SetDocument
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		private void PS_SD021_SetDocument(string oFormDocEntry)
		{
			try
			{
				if (string.IsNullOrEmpty(oFormDocEntry))
				{
					PS_SD021_FormItemEnabled();
					PS_SD021_AddMatrixRow(0, true); 
				}
				else
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD021_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_SD021_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				
				if (RowIserted == false) //행추가여부
				{
					oDS_PS_SD021L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_SD021L.Offset = oRow;
				oDS_PS_SD021L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		/// PS_SD021_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_SD021_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			int i;
			string sQry;
			string BPLID;
			string TeamCode;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "BPLID":
						BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();

						if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						//부서콤보세팅
						oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "전체");
						sQry = "  SELECT	U_Code AS [Code],";
						sQry += "           U_CodeNm As [Name]";
						sQry += " FROM      [@PS_HR200L]";
						sQry += " WHERE     Code = '1'";
						sQry += "           AND U_UseYN = 'Y'";
						sQry += "           AND U_Char2 = '" + BPLID + "'";
						sQry += " ORDER BY  U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode").Specific, sQry, "", false, false);
						oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);


						//매트릭스
						//구분
						sQry = "  SELECT	U_Minor AS [Code],";
						sQry += "           U_CdName As [Name]";
						sQry += " FROM      [@PS_SY001L]";
						sQry += " WHERE     Code = 'S006'";
						sQry += "           AND U_UseYN = 'Y'";
						sQry += "           AND U_RelCd = '" + BPLID + "'";
						sQry += " ORDER BY  U_Seq";

						if (oMat.Columns.Item("Class").ValidValues.Count > 0)
						{
							for (i = oMat.Columns.Item("Class").ValidValues.Count - 1; i >= 0; i += -1)
							{
								oMat.Columns.Item("Class").ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}
						dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("Class"), sQry,"","");
						break;

					case "TeamCode":
						TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();

						if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						//담당콤보세팅
						oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "전체");
						sQry = "  SELECT	U_Code AS [Code],";
						sQry += "           U_CodeNm As [Name]";
						sQry += " FROM      [@PS_HR200L]";
						sQry += " WHERE     Code = '2'";
						sQry += "           AND U_UseYN = 'Y'";
						sQry += "           AND U_Char1 = '" + TeamCode + "'";
						sQry += " ORDER BY  U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("RspCode").Specific, sQry, "", false, false);
						oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;

					case "CntcCode": //성명
						oForm.Items.Item("CntcName").Specific.Value = dataHelpClass.Get_ReData("U_FULLNAME", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("CntcCode").Specific.Value + "'","");
						break;

					case "Mat01":
						oMat.FlushToDataSource();

						//모든 컬럼이 입력되었을 때
						if (oMat.RowCount == oRow && !string.IsNullOrEmpty(oDS_PS_SD021L.GetValue("U_Class", oRow - 1).ToString().Trim()) 
							&& Convert.ToDouble(oDS_PS_SD021L.GetValue("U_Month01", oRow - 1).ToString().Trim()) != 0 
							&& Convert.ToDouble(oDS_PS_SD021L.GetValue("U_Month02", oRow - 1).ToString().Trim()) != 0
							&& Convert.ToDouble(oDS_PS_SD021L.GetValue("U_Month03", oRow - 1).ToString().Trim()) != 0
							&& Convert.ToDouble(oDS_PS_SD021L.GetValue("U_Month04", oRow - 1).ToString().Trim()) != 0
							&& Convert.ToDouble(oDS_PS_SD021L.GetValue("U_Month05", oRow - 1).ToString().Trim()) != 0
							&& Convert.ToDouble(oDS_PS_SD021L.GetValue("U_Month06", oRow - 1).ToString().Trim()) != 0
							&& Convert.ToDouble(oDS_PS_SD021L.GetValue("U_Month07", oRow - 1).ToString().Trim()) != 0
							&& Convert.ToDouble(oDS_PS_SD021L.GetValue("U_Month08", oRow - 1).ToString().Trim()) != 0
							&& Convert.ToDouble(oDS_PS_SD021L.GetValue("U_Month09", oRow - 1).ToString().Trim()) != 0
							&& Convert.ToDouble(oDS_PS_SD021L.GetValue("U_Month10", oRow - 1).ToString().Trim()) != 0
							&& Convert.ToDouble(oDS_PS_SD021L.GetValue("U_Month11", oRow - 1).ToString().Trim()) != 0 
							&& Convert.ToDouble(oDS_PS_SD021L.GetValue("U_Month12", oRow - 1).ToString().Trim()) != 0)
						{
							PS_SD021_AddMatrixRow(oRow, false); //행 추가
						}
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD021_FormItemEnabled
		/// </summary>
		private void PS_SD021_FormItemEnabled()
		{
			int i;
			string sQry;
			string BPLID;
			string TeamCode;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				//각 모드에 따른 아이템설정
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Mat01").Enabled = true;
					PS_SD021_FormClear();            
					oForm.EnableMenu("1281", true);	 //찾기
					oForm.EnableMenu("1282", false); //추가

					//사업장 선택
					oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

					//기준년도
					oForm.Items.Item("StdYear").Specific.Value = DateTime.Now.ToString("yyyy");
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("DocEntry").Specific.Value = "";
					oForm.Items.Item("DocEntry").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = false;
					oForm.EnableMenu("1281", false); //찾기
					oForm.EnableMenu("1282", true);	 //추가
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Mat01").Enabled = true;

					//사엽장예따른 콤보셋팅
					BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
					//부서콤보세팅
					if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
					{
						for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
						{
							oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
						}
					}
					oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "전체");
					sQry = "  SELECT	U_Code AS [Code],";
					sQry += "           U_CodeNm As [Name]";
					sQry += " FROM      [@PS_HR200L]";
					sQry += " WHERE     Code = '1'";
					sQry += "           AND U_UseYN = 'Y'";
					sQry += "           AND U_Char2 = '" + BPLID + "'";
					sQry += " ORDER BY  U_Seq";
					dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode").Specific, sQry, "", false, false);

					//담당콤보세팅
					TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();

					if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
					{
						for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
						{
							oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
						}
					}
					oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "전체");
					sQry = "  SELECT	U_Code AS [Code],";
					sQry += "           U_CodeNm As [Name]";
					sQry += " FROM      [@PS_HR200L]";
					sQry += " WHERE     Code = '2'";
					sQry += "           AND U_UseYN = 'Y'";
					sQry += "           AND U_Char1 = '" + TeamCode + "'";
					sQry += " ORDER BY  U_Seq";
					dataHelpClass.Set_ComboList(oForm.Items.Item("RspCode").Specific, sQry, "", false, false);
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
		/// PS_SD021_FormClear
		/// </summary>
		private void PS_SD021_FormClear()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			
			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD021'", "");
				if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
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
		}

		/// <summary>
		/// PS_SD021_DataValidCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_SD021_DataValidCheck()
		{
			bool functionReturnValue = false;
			int i;
			string errMessage = string.Empty;

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_SD021_FormClear();
				}

				if (string.IsNullOrEmpty(oForm.Items.Item("StdYear").Specific.Value.ToString().Trim()))
				{
					errMessage = "기준년도를 입력하지 않았습니다.";
					throw new Exception();
				}

				if (string.IsNullOrEmpty(oForm.Items.Item("BPLID").Specific.Selected.Value.ToString().Trim()))
				{
					errMessage = "사업장이 선택되지 않았습니다.";
					throw new Exception();
				}

				if (oForm.Items.Item("TeamCode").Specific.Selected.Value.ToString().Trim() == "%")
				{
					errMessage = "팀이 선택되지 않았습니다.";
					throw new Exception();
				}

				if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "사번을 입력하지 않았습니다.";
					throw new Exception();
				}

				if (oMat.VisualRowCount == 1)
				{
					errMessage = "라인이 존재하지 않습니다.";
					throw new Exception();
				}

				for (i = 1; i <= oMat.VisualRowCount - 1; i++)
				{
					if (string.IsNullOrEmpty(oMat.Columns.Item("Class").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("Class").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "구분이 선택되지 않았습니다.";
						throw new Exception();
					}
				}

				oDS_PS_SD021L.RemoveRecord(oDS_PS_SD021L.Size - 1);
				oMat.LoadFromDataSource();

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_SD021_FormClear();
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
				}
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_SD021_DataExistCheck 동일년도 동일 담당자의 2중 등록을 방지하기 위한 체크
		/// </summary>
		/// <param name="pMode"></param>
		/// <returns></returns>
		private bool PS_SD021_DataExistCheck(string pMode)
		{
			bool functionReturnValue = false;
			string sQry;
			string errMessage = string.Empty;

			int DocEntry;
			int CurDocEntry = 0;
			string StdYear;
			string CntcCode;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				StdYear = oForm.Items.Item("StdYear").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();

				if (pMode == "U")
				{
					CurDocEntry = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim());
				}

				sQry = "      SELECT   DocEntry";
				sQry += " FROM     [@PS_SD021H]";
				sQry += " WHERE   U_StdYear = '" + StdYear + "'";
				sQry += "             AND U_CntcCode = '" + CntcCode + "'";
				if (pMode == "U")
				{
					sQry += "         AND DocEntry <> " + CurDocEntry;
				}
				oRecordSet.DoQuery(sQry);

				if (oRecordSet.RecordCount != 0)
				{
					DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim());
					errMessage = "동일년도와 동일담당자로 등록된 데이터가 존재합니다. 확인하세요. 문서번호 : " + DocEntry;
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
				}
			}
			finally
            {
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_SD021_DataExistCheck("A") == false)
							{
								BubbleEvent = false;
								return;
							}

							if (PS_SD021_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							//해야할일 작업
							oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
							oFormMode01 = oForm.Mode;
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_SD021_DataExistCheck("U") == false)
							{
								BubbleEvent = false;
								return;
							}

							if (PS_SD021_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							//해야할일 작업
							oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
							oFormMode01 = oForm.Mode;
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_SD021_FormItemEnabled();
								PS_SD021_AddMatrixRow(0, true); 
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_SD021_FormItemEnabled();
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");
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
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_SD021_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
							oLastItemUID01 = pVal.ItemUID;
							oLastColUID01 = pVal.ColUID;
							oLastColRow01 = pVal.Row;

							oMat.SelectRow(pVal.Row, true, false);
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
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_SD021_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					PS_SD021_FormItemEnabled();
					PS_SD021_AddMatrixRow(oMat.VisualRowCount, false); 
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD021H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD021L);
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
		private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			int i;

			try
			{
				if (oLastColRow01 > 0)
				{
					if (pVal.BeforeAction == true)
					{
					}
					else if (pVal.BeforeAction == false)
					{
						for (i = 1; i <= oMat.VisualRowCount; i++)
						{
							oMat.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
						}
						oMat.FlushToDataSource();
						oDS_PS_SD021L.RemoveRecord(oDS_PS_SD021L.Size - 1);
						oMat.LoadFromDataSource();
						if (oMat.RowCount == 0)
						{
							PS_SD021_AddMatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_SD021L.GetValue("U_BatchNum", oMat.RowCount - 1).ToString().Trim()))
							{
								PS_SD021_AddMatrixRow(oMat.RowCount, false);
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
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
						case "1281": //찾기
							PS_SD021_FormItemEnabled(); 
							oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1282": //추가
							PS_SD021_FormItemEnabled();     
							PS_SD021_AddMatrixRow(0, true); 
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							PS_SD021_FormItemEnabled();
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
