using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 장비별 계획금액 VS 실적금액 조회(세부정보)
	/// </summary>
	internal class PS_CO923 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid01;
		private SAPbouiCOM.Grid oGrid02;
		private SAPbouiCOM.Grid oGrid03;
		private SAPbouiCOM.Grid oGrid04;
		private SAPbouiCOM.Grid oGrid05;
		
		private SAPbouiCOM.DataTable oDS_PS_CO923L;
		private SAPbouiCOM.DataTable oDS_PS_CO923M;
		private SAPbouiCOM.DataTable oDS_PS_CO923N;
		private SAPbouiCOM.DataTable oDS_PS_CO923O;
		private SAPbouiCOM.DataTable oDS_PS_CO923P;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO923.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_CO923_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_CO923");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_CO923_CreateItems();
				PS_CO923_ComboBox_Setting();

				oForm.Items.Item("BtnPrt11").Visible = false; //출력 버튼 비활성(2018.03.12 송명규) report없슴
				oForm.Items.Item("ItemCode").Click();
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
		/// PS_CO923_CreateItems
		/// </summary>
		private void PS_CO923_CreateItems()
		{
			try
			{
				oGrid01 = oForm.Items.Item("Grid01").Specific;
				oGrid02 = oForm.Items.Item("Grid02").Specific;
				oGrid03 = oForm.Items.Item("Grid03").Specific;
				oGrid04 = oForm.Items.Item("Grid04").Specific;
				oGrid05 = oForm.Items.Item("Grid05").Specific;

				oGrid01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
				oGrid02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
				oGrid03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
				oGrid04.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
				oGrid05.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

				oForm.DataSources.DataTables.Add("PS_CO923L");
				oForm.DataSources.DataTables.Add("PS_CO923M");
				oForm.DataSources.DataTables.Add("PS_CO923N");
				oForm.DataSources.DataTables.Add("PS_CO923O");
				oForm.DataSources.DataTables.Add("PS_CO923P");

				oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_CO923L");
				oGrid02.DataTable = oForm.DataSources.DataTables.Item("PS_CO923M");
				oGrid03.DataTable = oForm.DataSources.DataTables.Item("PS_CO923N");
				oGrid04.DataTable = oForm.DataSources.DataTables.Item("PS_CO923O");
				oGrid05.DataTable = oForm.DataSources.DataTables.Item("PS_CO923P");

				oDS_PS_CO923L = oForm.DataSources.DataTables.Item("PS_CO923L");
				oDS_PS_CO923M = oForm.DataSources.DataTables.Item("PS_CO923M");
				oDS_PS_CO923N = oForm.DataSources.DataTables.Item("PS_CO923N");
				oDS_PS_CO923O = oForm.DataSources.DataTables.Item("PS_CO923O");
				oDS_PS_CO923P = oForm.DataSources.DataTables.Item("PS_CO923P");

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID").Specific.DataBind.SetBound(true, "", "BPLID");

				//년월(시작)
				oForm.DataSources.UserDataSources.Add("FrMt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("FrMt").Specific.DataBind.SetBound(true, "", "FrMt");

				//년월(종료)
				oForm.DataSources.UserDataSources.Add("ToMt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("ToMt").Specific.DataBind.SetBound(true, "", "ToMt");

				//작번
				oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

				//품명
				oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

				//규격
				oForm.DataSources.UserDataSources.Add("ItemSpec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemSpec").Specific.DataBind.SetBound(true, "", "ItemSpec");

				oForm.Items.Item("FrMt").Specific.VALUE = DateTime.Now.ToString("yyyyMM");
				oForm.Items.Item("ToMt").Specific.VALUE = DateTime.Now.ToString("yyyyMM");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_CO923_ComboBox_Setting
		/// </summary>
		private void PS_CO923_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", dataHelpClass.User_BPLID(), false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_CO923_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_CO923_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "ItemCode":
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
		/// PS_CO923_FormResize
		/// </summary>
		private void PS_CO923_FormResize()
		{
			try
			{
				if (oGrid01.Columns.Count > 0)
				{
					oGrid01.AutoResizeColumns();
				}

				if (oGrid02.Columns.Count > 0)
				{
					oGrid02.AutoResizeColumns();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_CO923_MTX01 데이터 조회(집계)
		/// </summary>
		private void PS_CO923_MTX01()
		{
			string sQry;
			string errMessage = string.Empty;

			string CntcCode;
			string BPLID;
			string FrMt;
			string ToMt;
			string ItemCode;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				CntcCode = dataHelpClass.User_MSTCOD();
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				FrMt = oForm.Items.Item("FrMt").Specific.Value.ToString().Trim();
				ToMt = oForm.Items.Item("ToMt").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회 중(MTX01)...";

				sQry = " EXEC PS_CO923_01 '";
				sQry += CntcCode + "','";
				sQry += BPLID + "','";
				sQry += FrMt + "','";
				sQry += ToMt + "','";
				sQry += ItemCode + "'";

				oGrid01.DataTable.Clear();
				oDS_PS_CO923L.ExecuteQuery(sQry);

				oGrid01.Columns.Item(8).RightJustified = true;
				oGrid01.Columns.Item(10).RightJustified = true;
				oGrid01.Columns.Item(11).RightJustified = true;
				oGrid01.Columns.Item(14).RightJustified = true;
				oGrid01.Columns.Item(15).RightJustified = true;
				oGrid01.Columns.Item(17).RightJustified = true;
				oGrid01.Columns.Item(18).RightJustified = true;
				oGrid01.Columns.Item(20).RightJustified = true;
				oGrid01.Columns.Item(21).RightJustified = true;
				oGrid01.Columns.Item(22).RightJustified = true;
				oGrid01.Columns.Item(23).RightJustified = true;
				oGrid01.Columns.Item(24).RightJustified = true;
				oGrid01.Columns.Item(25).RightJustified = true;
				oGrid01.Columns.Item(26).RightJustified = true;

				if (oGrid01.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.(MTX01)";
					throw new Exception();
				}
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
				oGrid01.AutoResizeColumns();
				oForm.Update();
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_CO923_MTX02 데이터 조회(설계비)
		/// </summary>
		private void PS_CO923_MTX02()
		{
			string sQry;
			string errMessage = string.Empty;

			string CntcCode;
			string BPLID;
			string FrMt;
			string ToMt;
			string ItemCode;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				CntcCode = dataHelpClass.User_MSTCOD();
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				FrMt = oForm.Items.Item("FrMt").Specific.Value.ToString().Trim();
				ToMt = oForm.Items.Item("ToMt").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회 중(MTX02)...";

				sQry = " EXEC PS_CO923_02 '";
				sQry += CntcCode + "','";
				sQry += BPLID + "','";
				sQry += FrMt + "','";
				sQry += ToMt + "','";
				sQry += ItemCode + "'";

				oGrid02.DataTable.Clear();
				oDS_PS_CO923M.ExecuteQuery(sQry);

				oGrid02.Columns.Item(6).RightJustified = true;
				oGrid02.Columns.Item(7).RightJustified = true;
				oGrid02.Columns.Item(8).RightJustified = true;
				oGrid02.Columns.Item(9).RightJustified = true;

				if (oGrid02.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.(MTX02)";
					throw new Exception();
				}
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
				oGrid02.AutoResizeColumns();
				oForm.Update();
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_CO923_MTX03 데이터 조회(원재료비)
		/// </summary>
		private void PS_CO923_MTX03()
		{
			string sQry;
			string errMessage = string.Empty;

			string BPLID;
			string FrMt;
			string ToMt;
			string ItemCode;
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				FrMt = oForm.Items.Item("FrMt").Specific.Value.ToString().Trim();
				ToMt = oForm.Items.Item("ToMt").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회 중(MTX03)...";

				sQry = " EXEC PS_CO923_03 '";
				sQry += BPLID + "','";
				sQry += FrMt + "','";
				sQry += ToMt + "','";
				sQry += ItemCode + "'";

				oGrid03.DataTable.Clear();
				oDS_PS_CO923N.ExecuteQuery(sQry);

				oGrid03.Columns.Item(3).RightJustified = true;

				if (oGrid03.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.(MTX03)";
					throw new Exception();
				}
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
				oGrid03.AutoResizeColumns();
				oForm.Update();
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_CO923_MTX04 데이터 조회(가공비)
		/// </summary>
		private void PS_CO923_MTX04()
		{
			string sQry;
			string errMessage = string.Empty;

			string CntcCode;
			string BPLID;
			string FrMt;
			string ToMt;
			string ItemCode;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				CntcCode = dataHelpClass.User_MSTCOD();
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				FrMt = oForm.Items.Item("FrMt").Specific.Value.ToString().Trim();
				ToMt = oForm.Items.Item("ToMt").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회 중(MTX04)...";

				sQry = " EXEC PS_CO923_04 '";
				sQry += CntcCode + "','";
				sQry += BPLID + "','";
				sQry += FrMt + "','";
				sQry += ToMt + "','";
				sQry += ItemCode + "'";

				oGrid04.DataTable.Clear();
				oDS_PS_CO923O.ExecuteQuery(sQry);

				oGrid04.Columns.Item(6).RightJustified = true;
				oGrid04.Columns.Item(7).RightJustified = true;
				oGrid04.Columns.Item(8).RightJustified = true;
				oGrid04.Columns.Item(9).RightJustified = true;

				if (oGrid04.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.(MTX04)";
					throw new Exception();
				}
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
				oGrid04.AutoResizeColumns();
				oForm.Update();
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_CO923_MTX05 데이터 조회(외주제작,외주가공)
		/// </summary>
		private void PS_CO923_MTX05()
		{
			string sQry;
			string errMessage = string.Empty;

			string BPLID;
			string FrMt;
			string ToMt;
			string ItemCode;
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				FrMt = oForm.Items.Item("FrMt").Specific.Value.ToString().Trim();
				ToMt = oForm.Items.Item("ToMt").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회 중(MTX05)...";

				sQry = "EXEC PS_CO923_05 '";
				sQry += BPLID + "','";
				sQry += FrMt + "','";
				sQry += ToMt + "','";
				sQry += ItemCode + "'";

				oGrid05.DataTable.Clear();
				oDS_PS_CO923P.ExecuteQuery(sQry);

				oGrid05.Columns.Item(5).RightJustified = true;
				oGrid05.Columns.Item(6).RightJustified = true;
				oGrid05.Columns.Item(7).RightJustified = true;

				if (oGrid05.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.(MTX05)";
					throw new Exception();
				}
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
				oGrid05.AutoResizeColumns();
				oForm.Update();
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
				PSH_Globals.SBO_Application.MessageBox("조회 완료.");
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
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "BtnSearch")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_CO923_MTX01();
							PS_CO923_MTX02();
							PS_CO923_MTX03();
							PS_CO923_MTX04();
							PS_CO923_MTX05();
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
					PS_CO923_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					if (pVal.ItemUID == "Grid01")
					{
					}
					else if (pVal.ItemUID == "lblGrid02")
					{
						if (oForm.Items.Item("Grid02").Width == 310)
						{
							oForm.Freeze(true);
							oForm.Items.Item("Grid02").Width = 1280; //나머지 그리드 안보이게 함
							oForm.Items.Item("Grid03").Visible = false;
							oForm.Items.Item("lblGrid03").Visible = false;
							oForm.Items.Item("Grid04").Visible = false;
							oForm.Items.Item("lblGrid04").Visible = false;
							oForm.Items.Item("Grid05").Visible = false;
							oForm.Items.Item("lblGrid05").Visible = false;
							oForm.Freeze(false);
						}
						else
						{
							oForm.Freeze(true);
							oForm.Items.Item("Grid02").Width = 310; //나머지 그리드 다시 보이게 함
							oForm.Items.Item("Grid03").Visible = true;
							oForm.Items.Item("lblGrid03").Visible = true;
							oForm.Items.Item("Grid04").Visible = true;
							oForm.Items.Item("lblGrid04").Visible = true;
							oForm.Items.Item("Grid05").Visible = true;
							oForm.Items.Item("lblGrid05").Visible = true;
							oForm.Freeze(false);
						}

						if (oGrid02.Columns.Count > 0)
						{
							oGrid02.AutoResizeColumns();
						}
					}
					else if (pVal.ItemUID == "lblGrid03")
					{
						if (oForm.Items.Item("Grid03").Width == 310)
						{
							oForm.Freeze(true);
							oForm.Items.Item("Grid03").Width = 1280;
							oForm.Items.Item("Grid03").Left = 10;
							oForm.Items.Item("lblGrid03").Left = 10; //나머지 그리드 안보이게 함
							oForm.Items.Item("Grid02").Visible = false;
							oForm.Items.Item("lblGrid02").Visible = false;
							oForm.Items.Item("Grid04").Visible = false;
							oForm.Items.Item("lblGrid04").Visible = false;
							oForm.Items.Item("Grid05").Visible = false;
							oForm.Items.Item("lblGrid05").Visible = false;
							oForm.Freeze(false);
						}
						else
						{
							oForm.Freeze(true);
							oForm.Items.Item("Grid03").Width = 310;
							oForm.Items.Item("Grid03").Left = 334;
							oForm.Items.Item("lblGrid03").Left = 334; //나머지 그리드 다시 보이게 함
							oForm.Items.Item("Grid02").Visible = true;
							oForm.Items.Item("lblGrid02").Visible = true;
							oForm.Items.Item("Grid04").Visible = true;
							oForm.Items.Item("lblGrid04").Visible = true;
							oForm.Items.Item("Grid05").Visible = true;
							oForm.Items.Item("lblGrid05").Visible = true;
							oForm.Freeze(false);
						}

						if (oGrid03.Columns.Count > 0)
						{
							oGrid03.AutoResizeColumns();
						}
					}
					else if (pVal.ItemUID == "lblGrid04")
					{
						if (oForm.Items.Item("Grid04").Width == 310)
						{
							oForm.Freeze(true);
							oForm.Items.Item("Grid04").Width = 1280;
							oForm.Items.Item("Grid04").Left = 10;
							oForm.Items.Item("lblGrid04").Left = 10; //나머지 그리드 안보이게 함
							oForm.Items.Item("Grid02").Visible = false;
							oForm.Items.Item("lblGrid02").Visible = false;
							oForm.Items.Item("Grid03").Visible = false;
							oForm.Items.Item("lblGrid03").Visible = false;
							oForm.Items.Item("Grid05").Visible = false;
							oForm.Items.Item("lblGrid05").Visible = false;
							oForm.Freeze(false);
						}
						else
						{
							oForm.Freeze(true);
							oForm.Items.Item("Grid04").Width = 310;
							oForm.Items.Item("Grid04").Left = 657;
							oForm.Items.Item("lblGrid04").Left = 657; //나머지 그리드 다시 보이게 함
							oForm.Items.Item("Grid02").Visible = true;
							oForm.Items.Item("lblGrid02").Visible = true;
							oForm.Items.Item("Grid03").Visible = true;
							oForm.Items.Item("lblGrid03").Visible = true;
							oForm.Items.Item("Grid05").Visible = true;
							oForm.Items.Item("lblGrid05").Visible = true;
							oForm.Freeze(false);
						}

						if (oGrid04.Columns.Count > 0)
						{
							oGrid04.AutoResizeColumns();
						}
					}
					else if (pVal.ItemUID == "lblGrid05")
					{
						if (oForm.Items.Item("Grid05").Width == 310)
						{
							oForm.Freeze(true);
							oForm.Items.Item("Grid05").Width = 1280;
							oForm.Items.Item("Grid05").Left = 10;
							oForm.Items.Item("lblGrid05").Left = 10; //나머지 그리드 안보이게 함
							oForm.Items.Item("Grid02").Visible = false;
							oForm.Items.Item("lblGrid02").Visible = false;
							oForm.Items.Item("Grid03").Visible = false;
							oForm.Items.Item("lblGrid03").Visible = false;
							oForm.Items.Item("Grid04").Visible = false;
							oForm.Items.Item("lblGrid04").Visible = false;
							oForm.Freeze(false);
						}
						else
						{
							oForm.Freeze(true);
							oForm.Items.Item("Grid05").Width = 310;
							oForm.Items.Item("Grid05").Left = 980;
							oForm.Items.Item("lblGrid05").Left = 980; //나머지 그리드 다시 보이게 함
							oForm.Items.Item("Grid02").Visible = true;
							oForm.Items.Item("lblGrid02").Visible = true;
							oForm.Items.Item("Grid03").Visible = true;
							oForm.Items.Item("lblGrid03").Visible = true;
							oForm.Items.Item("Grid04").Visible = true;
							oForm.Items.Item("lblGrid04").Visible = true;
							oForm.Freeze(false);
						}

						if (oGrid05.Columns.Count > 0)
						{
							oGrid05.AutoResizeColumns();
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
					PS_CO923_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					PS_CO923_FormResize();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid03);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid04);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid05);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO923L);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO923M);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO923N);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO923O);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO923P);
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
