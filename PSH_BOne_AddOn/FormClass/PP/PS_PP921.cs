using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 설계비 배부 결과 조회
	/// </summary>
	internal class PS_PP921 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid01;
		private SAPbouiCOM.Grid oGrid02;
		private SAPbouiCOM.Grid oGrid03;
		private SAPbouiCOM.Grid oGrid04;

		private SAPbouiCOM.DataTable oDS_PS_PP921A;
		private SAPbouiCOM.DataTable oDS_PS_PP921B;
		private SAPbouiCOM.DataTable oDS_PS_PP921C;
		private SAPbouiCOM.DataTable oDS_PS_PP921D;

		/// <summary>
		/// 화면 호출
		/// </summary>
		public override void LoadForm(string oFormDocEntry01)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP921.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP921_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP921");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP921_CreateItems();
				PS_PP921_ComboBox_Setting();
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
				oForm.Items.Item("Folder01").Specific.Select(); //폼이 로드 될 때 Folder01이 선택됨
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_PP921_CreateItems
		/// </summary>
		private void PS_PP921_CreateItems()
		{
			try
			{
				oGrid01 = oForm.Items.Item("Grid01").Specific;
				oGrid02 = oForm.Items.Item("Grid02").Specific;
				oGrid03 = oForm.Items.Item("Grid03").Specific;
				oGrid04 = oForm.Items.Item("Grid04").Specific;

				oForm.DataSources.DataTables.Add("PS_PP921A");
				oForm.DataSources.DataTables.Add("PS_PP921B");
				oForm.DataSources.DataTables.Add("PS_PP921C");
				oForm.DataSources.DataTables.Add("PS_PP921D");

				oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_PP921A");
				oGrid02.DataTable = oForm.DataSources.DataTables.Item("PS_PP921B");
				oGrid03.DataTable = oForm.DataSources.DataTables.Item("PS_PP921C");
				oGrid04.DataTable = oForm.DataSources.DataTables.Item("PS_PP921D");

				oDS_PS_PP921A = oForm.DataSources.DataTables.Item("PS_PP921A");
				oDS_PS_PP921B = oForm.DataSources.DataTables.Item("PS_PP921B");
				oDS_PS_PP921C = oForm.DataSources.DataTables.Item("PS_PP921C");
				oDS_PS_PP921D = oForm.DataSources.DataTables.Item("PS_PP921D");

				//기준년월
				oForm.DataSources.UserDataSources.Add("StdYM01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("StdYM01").Specific.DataBind.SetBound(true, "", "StdYM01");

				//구분
				oForm.DataSources.UserDataSources.Add("Class01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Class01").Specific.DataBind.SetBound(true, "", "Class01");

				//기준년월
				oForm.DataSources.UserDataSources.Add("StdYM02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("StdYM02").Specific.DataBind.SetBound(true, "", "StdYM02");

				//기준년월
				oForm.DataSources.UserDataSources.Add("StdYM03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("StdYM03").Specific.DataBind.SetBound(true, "", "StdYM03");

				//기준년월
				oForm.DataSources.UserDataSources.Add("StdYM04", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("StdYM04").Specific.DataBind.SetBound(true, "", "StdYM04");

				oForm.DataSources.UserDataSources.Item("StdYM01").Value = DateTime.Now.AddMonths(-1).ToString("yyyyMM"); //전월
				oForm.DataSources.UserDataSources.Item("StdYM02").Value = DateTime.Now.AddMonths(-1).ToString("yyyyMM"); //전월
				oForm.DataSources.UserDataSources.Item("StdYM03").Value = DateTime.Now.AddMonths(-1).ToString("yyyyMM"); //전월
				oForm.DataSources.UserDataSources.Item("StdYM04").Value = DateTime.Now.AddMonths(-1).ToString("yyyyMM"); //전월
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP921_ComboBox_Setting
		/// </summary>
		private void PS_PP921_ComboBox_Setting()
		{
			try
			{
				oForm.Items.Item("Class01").Specific.ValidValues.Add("01", "담당별");
				oForm.Items.Item("Class01").Specific.ValidValues.Add("02", "개인별");
				oForm.Items.Item("Class01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP921_MTX01  작번별 공수 조회
		/// </summary>
		private void PS_PP921_MTX01()
		{
			string sQry;
			string errMessage = String.Empty;

			string StdYM;		  //기준년월
			string Class_Renamed; //구분

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				StdYM = oForm.Items.Item("StdYM01").Specific.Value.ToString().Trim();
				Class_Renamed = oForm.Items.Item("Class01").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				if (Class_Renamed == "01") //담당별
				{
					sQry = " EXEC PS_PP921_01 '";
					sQry += StdYM + "'";
				}
				else //개인별
				{
					sQry = " EXEC PS_PP921_02 '";
					sQry += StdYM + "'";
				}

				oGrid01.DataTable.Clear();
				oDS_PS_PP921A.ExecuteQuery(sQry);
				
				if (Class_Renamed == "01") //담당별
				{
					oGrid01.Columns.Item(4).RightJustified = true;
					oGrid01.Columns.Item(5).RightJustified = true;
					oGrid01.Columns.Item(6).RightJustified = true;
					oGrid01.Columns.Item(7).RightJustified = true;
				}
				else //개인별
				{
					oGrid01.Columns.Item(11).RightJustified = true;
				}

				if (oGrid01.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
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
		/// PS_PP921_MTX02  비용 배부 결과 조회
		/// </summary>
		private void PS_PP921_MTX02()
		{
			string sQry;
			string errMessage = String.Empty;

			string StdYM;         //기준년월

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				StdYM = oForm.Items.Item("StdYM02").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = " EXEC PS_PP921_11 '";
				sQry += StdYM + "'";

				oGrid02.DataTable.Clear();
				oDS_PS_PP921B.ExecuteQuery(sQry);

				oGrid02.Columns.Item(4).RightJustified = true;

				if (oGrid02.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
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
		/// PS_PP921_MTX03  작번별 배부 결과
		/// </summary>
		private void PS_PP921_MTX03()
		{
			string sQry;
			string errMessage = String.Empty;

			string StdYM;         //기준년월

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				StdYM = oForm.Items.Item("StdYM03").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = " EXEC PS_PP921_21 '";
				sQry += StdYM + "'";

				oGrid03.DataTable.Clear();
				oDS_PS_PP921C.ExecuteQuery(sQry);

				oGrid03.Columns.Item(5).RightJustified = true;
				oGrid03.Columns.Item(6).RightJustified = true;
				oGrid03.Columns.Item(7).RightJustified = true;
				oGrid03.Columns.Item(8).RightJustified = true;
				oGrid03.Columns.Item(9).RightJustified = true;
				oGrid03.Columns.Item(10).RightJustified = true;
				oGrid03.Columns.Item(11).RightJustified = true;
				oGrid03.Columns.Item(12).RightJustified = true;
				oGrid03.Columns.Item(13).RightJustified = true;
				oGrid03.Columns.Item(14).RightJustified = true;
				oGrid03.Columns.Item(15).RightJustified = true;
				oGrid03.Columns.Item(16).RightJustified = true;
				oGrid03.Columns.Item(17).RightJustified = true;

				if (oGrid03.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
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
		/// PS_PP921_MTX04  작번별 배부 결과(누계)
		/// </summary>
		private void PS_PP921_MTX04()
		{
			string sQry;
			string errMessage = String.Empty;

			string StdYM;         //기준년월

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				StdYM = oForm.Items.Item("StdYM04").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = " EXEC PS_PP921_31 '";
				sQry += StdYM + "'";

				oGrid04.DataTable.Clear();
				oDS_PS_PP921D.ExecuteQuery(sQry);

				oGrid04.Columns.Item(4).RightJustified = true;
				oGrid04.Columns.Item(5).RightJustified = true;
				oGrid04.Columns.Item(6).RightJustified = true;
				oGrid04.Columns.Item(7).RightJustified = true;
				oGrid04.Columns.Item(8).RightJustified = true;
				oGrid04.Columns.Item(9).RightJustified = true;

				if (oGrid04.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
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
		/// PS_PP921_Print_Report01  비용 배부 결과 출력
		/// </summary>
		[STAThread]
		private void PS_PP921_Print_Report01()
		{
			string WinTitle;
			string ReportName;

			string StdYM;         //기준년월
			string Class_Renamed; //구분

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				StdYM = oForm.Items.Item("StdYM01").Specific.Value.ToString().Trim();
				Class_Renamed = oForm.Items.Item("Class01").Specific.Value.ToString().Trim();

				WinTitle = "[PS_PP921] 레포트";
				
				if (Class_Renamed == "01") //담당별
				{
					ReportName = "PS_PP921_01.rpt";
				}
				else  //개인별
				{
					ReportName = "PS_PP921_02.rpt";
				}

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@StdYM", StdYM));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP921_Print_Report02  비용 배부 결과 출력
		/// </summary>
		[STAThread]
		private void PS_PP921_Print_Report02()
		{
			string WinTitle;
			string ReportName;
			string StdYM;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				StdYM = oForm.Items.Item("StdYM02").Specific.Value.ToString().Trim();

				WinTitle = "[PS_PP921] 레포트";
				ReportName = "PS_PP921_03.rpt";

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@StdYM", StdYM));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP921_Print_Report03  작번별 배부 결과 출력
		/// </summary>
		[STAThread]
		private void PS_PP921_Print_Report03()
		{
			string WinTitle;
			string ReportName;
			string StdYM;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				StdYM = oForm.Items.Item("StdYM03").Specific.Value.ToString().Trim();

			    WinTitle = "[PS_PP921] 레포트";
			    ReportName = "PS_PP921_04.rpt";

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@StdYM", StdYM));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP921_Print_Report04  작번별 배부 결과 출력(누계)
		/// </summary>
		[STAThread]
		private void PS_PP921_Print_Report04()
		{

			string WinTitle;
			string ReportName;
			string StdYM;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				StdYM = oForm.Items.Item("StdYM04").Specific.Value.ToString().Trim();

			    WinTitle = "[PS_PP921] 레포트";
			    ReportName = "PS_PP921_05.rpt";

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@StdYM", StdYM));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
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
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
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
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_PP921_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnSrch02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_PP921_MTX02();
						}
					}
					else if (pVal.ItemUID == "BtnSrch03")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_PP921_MTX03();
						}
					}
					else if (pVal.ItemUID == "BtnSrch04")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_PP921_MTX04();
						}
					}
					else if (pVal.ItemUID == "BtnPrt01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_PP921_Print_Report01);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
					else if (pVal.ItemUID == "BtnPrt02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_PP921_Print_Report02);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
					else if (pVal.ItemUID == "BtnPrt03")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_PP921_Print_Report03);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
					else if (pVal.ItemUID == "BtnPrt04")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_PP921_Print_Report04);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					//폴더를 사용할 때는 필수 소스_S
					if (pVal.ItemUID == "Folder01")
					{
						oForm.PaneLevel = 1;
						oForm.DefButton = "BtnSrch01";
					}
					if (pVal.ItemUID == "Folder02")
					{
						oForm.PaneLevel = 2;
						oForm.DefButton = "BtnSrch02";
					}
					if (pVal.ItemUID == "Folder03")
					{
						oForm.PaneLevel = 3;
						oForm.DefButton = "BtnSrch03";
					}
					if (pVal.ItemUID == "Folder04")
					{
						oForm.PaneLevel = 4;
						oForm.DefButton = "BtnSrch04";
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					//그룹박스 크기 동적 할당
					oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Grid01").Height + 80;
					oForm.Items.Item("GrpBox01").Width = oForm.Items.Item("Grid01").Width + 30;

					if (oGrid01.Columns.Count > 0)
					{
						oGrid01.AutoResizeColumns();
					}

					if (oGrid02.Columns.Count > 0)
					{
						oGrid02.AutoResizeColumns();
					}

					if (oGrid03.Columns.Count > 0)
					{
						oGrid03.AutoResizeColumns();
					}
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
			if (pVal.Before_Action == true)
			{
			}
			else if (pVal.Before_Action == false)
			{
				SubMain.Remove_Forms(oFormUniqueID);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid02);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid03);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid04);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP921A);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP921B);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP921C);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP921D);
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
