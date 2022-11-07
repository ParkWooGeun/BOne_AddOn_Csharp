using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 검사일 조회
	/// </summary>
	internal class PS_QM131 : PSH_BaseClass
	{
		private string oFormUniqueID;

		private SAPbouiCOM.Grid oGrid11;
		private SAPbouiCOM.Grid oGrid12;
		private SAPbouiCOM.DataTable oDS_PS_QM131A;
		private SAPbouiCOM.DataTable oDS_PS_QM131B;

		private SAPbouiCOM.Grid oGrid21;
		private SAPbouiCOM.Grid oGrid22;
		private SAPbouiCOM.DataTable oDS_PS_QM131C;
		private SAPbouiCOM.DataTable oDS_PS_QM131D;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM131.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM131_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM131");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_QM131_CreateItems();
				PS_QM131_ComboBox_Setting();
				oForm.Items.Item("Folder01").Specific.Select(); //폼이 로드 될 때 Folder01이 선택됨
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
		/// PS_QM131_CreateItems
		/// </summary>
		private void PS_QM131_CreateItems()
		{
			try
			{
				oGrid11 = oForm.Items.Item("Grid11").Specific;
				oGrid12 = oForm.Items.Item("Grid12").Specific;
				oGrid21 = oForm.Items.Item("Grid21").Specific;
				oGrid22 = oForm.Items.Item("Grid22").Specific;

				oGrid11.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
				oGrid12.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
				oGrid21.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
				oGrid22.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

				oForm.DataSources.DataTables.Add("PS_QM131A");
				oForm.DataSources.DataTables.Add("PS_QM131B");
				oForm.DataSources.DataTables.Add("PS_QM131C");
				oForm.DataSources.DataTables.Add("PS_QM131D");

				oGrid11.DataTable = oForm.DataSources.DataTables.Item("PS_QM131A");
				oGrid12.DataTable = oForm.DataSources.DataTables.Item("PS_QM131B");
				oGrid21.DataTable = oForm.DataSources.DataTables.Item("PS_QM131C");
				oGrid22.DataTable = oForm.DataSources.DataTables.Item("PS_QM131D");

				oDS_PS_QM131A = oForm.DataSources.DataTables.Item("PS_QM131A");
				oDS_PS_QM131B = oForm.DataSources.DataTables.Item("PS_QM131B");
				oDS_PS_QM131C = oForm.DataSources.DataTables.Item("PS_QM131C");
				oDS_PS_QM131D = oForm.DataSources.DataTables.Item("PS_QM131D");

				//출고일
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID11", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID11").Specific.DataBind.SetBound(true, "", "BPLID11");

				//반출기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDt11", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt11").Specific.DataBind.SetBound(true, "", "FrDt11");

				//반출기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDt11", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt11").Specific.DataBind.SetBound(true, "", "ToDt11");

				//생산완료일
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID21", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID21").Specific.DataBind.SetBound(true, "", "BPLID21");

				//가입고기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDt21", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt21").Specific.DataBind.SetBound(true, "", "FrDt21");

				//가입고기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDt21", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt21").Specific.DataBind.SetBound(true, "", "ToDt21");

				//자체/외주
				oForm.DataSources.UserDataSources.Add("InOut21", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("InOut21").Specific.DataBind.SetBound(true, "", "InOut21");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType21", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("CardType21").Specific.DataBind.SetBound(true, "", "CardType21");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemType21", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("ItemType21").Specific.DataBind.SetBound(true, "", "ItemType21");

				//품목코드
				oForm.DataSources.UserDataSources.Add("ItemCode21", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("ItemCode21").Specific.DataBind.SetBound(true, "", "ItemCode21");

				//품목명
				oForm.DataSources.UserDataSources.Add("ItemName21", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemName21").Specific.DataBind.SetBound(true, "", "ItemName21");

				oForm.Items.Item("FrDt11").Specific.Value = DateTime.Now.ToString("yyyyMM01");
				oForm.Items.Item("ToDt11").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("FrDt21").Specific.Value = DateTime.Now.AddMonths(-2).ToString("yyyyMM") + "01";
				oForm.Items.Item("ToDt21").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM131_ComboBox_Setting
		/// </summary>
		private void PS_QM131_ComboBox_Setting()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장(출고일)
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID11").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", dataHelpClass.User_BPLID(), false, false);

				//사업장(생산완료일)
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID21").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", dataHelpClass.User_BPLID(), false, false);

				//자체/외주(생산완료일)
				oForm.Items.Item("InOut21").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("InOut21").Specific.ValidValues.Add("IN", "자체");
				oForm.Items.Item("InOut21").Specific.ValidValues.Add("OUT", "외주");
				oForm.Items.Item("InOut21").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//거래처구분(생산완료일)
				sQry = "     SELECT  U_Minor,";
				sQry += "             U_CdName";
				sQry += "  FROM   [@PS_SY001L]";
				sQry += "  WHERE  Code = 'C100'";
				oForm.Items.Item("CardType21").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType21").Specific, sQry, "%", false, false);

				//품목구분(생산완료일)
				sQry = "     SELECT  U_Minor,";
				sQry += "             U_CdName";
				sQry += "  FROM   [@PS_SY001L]";
				sQry += "  WHERE  Code = 'S002'";
				oForm.Items.Item("ItemType21").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType21").Specific, sQry, "%", false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM131_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		private void PS_QM131_FlushToItemValue(string oUID)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "ItemCode21":
						oForm.Items.Item("ItemName21").Specific.Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item("ItemCode21").Specific.Value.ToString().Trim() + "'", "");
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM131_MTX11
		/// </summary>
		private void PS_QM131_MTX11()
		{
			string BPLID;
			string FrDt;
			string ToDt;
			string sQry;
			string errMessage = string.Empty;
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				BPLID = oForm.Items.Item("BPLID11").Specific.Value;
				FrDt = oForm.Items.Item("FrDt11").Specific.Value;
				ToDt = oForm.Items.Item("ToDt11").Specific.Value;

				ProgressBar01.Text = "조회시작!";

				sQry = " EXEC PS_QM131_11 '";
				sQry += BPLID + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "'";

				oGrid11.DataTable.Clear();
				oDS_PS_QM131A.ExecuteQuery(sQry);

				oGrid11.Columns.Item(3).RightJustified = true;
				oGrid12.DataTable.Clear(); //상세 그리드도 클리어

				if (oGrid11.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid11.AutoResizeColumns();
				oForm.Update();
			}
			catch (Exception ex)
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
				}
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_QM131_MTX12
		/// </summary>
		private void PS_QM131_MTX12()
		{
			int loopCount1;
			string ItemCode = string.Empty; //자재코드
			string FrDt;
			string ToDt;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			string errMessage = string.Empty;
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				FrDt = oForm.Items.Item("FrDt11").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt11").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				for (loopCount1 = 0; loopCount1 <= oGrid11.Rows.Count - 1; loopCount1++)
				{
					if (oGrid11.Rows.IsSelected(loopCount1) == true)
					{
						ItemCode = oGrid11.DataTable.GetValue(0, loopCount1).ToString().Trim();
					}
				}

				sQry = "    EXEC PS_QM131_12 '";
				sQry += ItemCode + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "'";

				oGrid12.DataTable.Clear();
				oDS_PS_QM131B.ExecuteQuery(sQry);
				oRecordSet.DoQuery(sQry);

				if (oGrid12.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid12.AutoResizeColumns();
				oForm.Update();
			}
			catch (Exception ex)
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
				}
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
		/// PS_QM131_MTX21
		/// </summary>
		private void PS_QM131_MTX21()
		{
			string BPLID;
			string FrDt;
			string ToDt;
			string InOut;
			string CardType;
			string ItemType;
			string ItemCode;
			string sQry;
			string errMessage = string.Empty;
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				BPLID = oForm.Items.Item("BPLID21").Specific.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt21").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt21").Specific.Value.ToString().Trim();
				InOut = oForm.Items.Item("InOut21").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType21").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType21").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode21").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "    EXEC PS_QM131_21 '";
				sQry += BPLID + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += InOut + "','";
				sQry += CardType + "','";
				sQry += ItemType + "','";
				sQry += ItemCode + "'";

				oGrid21.DataTable.Clear();
				oDS_PS_QM131C.ExecuteQuery(sQry);

				oGrid21.Columns.Item(7).RightJustified = true;
				oGrid21.Columns.Item(10).RightJustified = true;

				oGrid22.DataTable.Clear(); //상세 그리드도 클리어

				if (oGrid21.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid21.AutoResizeColumns();
				oForm.Update();
			}
			catch (Exception ex)
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
				}
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_QM131_MTX22
		/// </summary>
		private void PS_QM131_MTX22()
		{
			int loopCount1;
			string ItemCode = string.Empty; //자재코드
			string FrDt;
			string ToDt;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			string errMessage = string.Empty;
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				FrDt = oForm.Items.Item("FrDt21").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt21").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				for (loopCount1 = 0; loopCount1 <= oGrid21.Rows.Count - 1; loopCount1++)
				{
					if (oGrid21.Rows.IsSelected(loopCount1) == true)
					{
						ItemCode = oGrid21.DataTable.GetValue(0, loopCount1).ToString().Trim();
					}

				}

				sQry = "    EXEC PS_QM131_22 '";
				sQry += ItemCode + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "'";

				oGrid22.DataTable.Clear();
				oDS_PS_QM131D.ExecuteQuery(sQry);

				oGrid22.Columns.Item(4).RightJustified = true;
				oGrid22.Columns.Item(5).RightJustified = true;
				oRecordSet.DoQuery(sQry);

				if (oGrid22.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid22.AutoResizeColumns();
				oForm.Update();
			}
			catch (Exception ex)
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
				}
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
		/// PS_QM131_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_QM131_Print_Report01()
		{
			string WinTitle;
			string ReportName;
			string BPLID;
			string FrDt;
			string ToDt;
			string InOut;
			string CardType;
			string ItemType;
			string ItemCode;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID = oForm.Items.Item("BPLID21").Specific.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt21").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt21").Specific.Value.ToString().Trim();
				InOut = oForm.Items.Item("InOut21").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType21").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType21").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode21").Specific.Value.ToString().Trim();

				WinTitle = "[PS_QM131] 레포트";
				ReportName = "PS_QM131_01.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@FrDt", DateTime.ParseExact(FrDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@ToDt", DateTime.ParseExact(ToDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@InOut", InOut));
				dataPackParameter.Add(new PSH_DataPackClass("@CardType", CardType));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemType", ItemType));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //	Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
					Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //	Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                //    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                //    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
					Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
					break;
			}
		}

		/// <summary>
		/// ITEM_PRESSED 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "BtnSrch11")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_QM131_MTX11();
						}
					}
					else if (pVal.ItemUID == "BtnSrch21")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_QM131_MTX21();
						}
					}
					else if (pVal.ItemUID == "BtnPrt21")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_QM131_Print_Report01);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Folder01") //Folder01이 선택되었을 때
					{
						oForm.PaneLevel = 1;
						oForm.DefButton = "BtnSrch11";
					}
					if (pVal.ItemUID == "Folder02") //Folder02가 선택되었을 때
					{
						oForm.PaneLevel = 2;
						oForm.DefButton = "BtnSrch21";
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// KEY_DOWN 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode21", "");
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
					if (pVal.ItemUID == "Grid11")
					{
						if (pVal.Row == -1)
						{
							oGrid11.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
						}
						else
						{
							if (oGrid11.Rows.SelectedRows.Count > 0)
							{
								PS_QM131_MTX12();
							}
						}
					}
					else if (pVal.ItemUID == "Grid21")
					{
						if (pVal.Row == -1)
						{
							oGrid21.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
						}
						else
						{
							if (oGrid21.Rows.SelectedRows.Count > 0)
							{
								PS_QM131_MTX22();
							}
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
					PS_QM131_FlushToItemValue(pVal.ItemUID);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
					oForm.Items.Item("Grid11").Height = oForm.Height - 170;
					oForm.Items.Item("Grid11").Width = oForm.Width / 2 + 430;
					oForm.Items.Item("Grid12").Left = oForm.Items.Item("Grid11").Width + 40;
					oForm.Items.Item("Grid12").Height = oForm.Items.Item("Grid11").Height;
					oForm.Items.Item("Grid12").Width = oForm.Width - oForm.Items.Item("Grid11").Width - 75;

					oForm.Items.Item("Grid21").Height = oForm.Items.Item("Grid11").Height - 45;
					oForm.Items.Item("Grid21").Width = oForm.Items.Item("Grid11").Width;
					oForm.Items.Item("Grid22").Left = oForm.Items.Item("Grid12").Left;
					oForm.Items.Item("Grid22").Height = oForm.Items.Item("Grid12").Height - 45;
					oForm.Items.Item("Grid22").Width = oForm.Items.Item("Grid12").Width;
					//그룹박스 크기 동적 할당
					oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Grid11").Height + 70;
					oForm.Items.Item("GrpBox01").Width = oForm.Items.Item("Grid11").Width + oForm.Items.Item("Grid12").Width + 45;

					if (oGrid11.Columns.Count > 0)
					{
						oGrid11.AutoResizeColumns();
					}
					if (oGrid12.Columns.Count > 0)
					{
						oGrid12.AutoResizeColumns();
					}
					if (oGrid21.Columns.Count > 0)
					{
						oGrid21.AutoResizeColumns();
					}
					if (oGrid22.Columns.Count > 0)
					{
						oGrid22.AutoResizeColumns();
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// VALIDATE 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
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
					PS_QM131_FlushToItemValue(pVal.ItemUID);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				oForm.Freeze(false);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid11);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid12);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid21);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid22);

					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM131A);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM131B);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM131C);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM131D);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
						case "1287": //복제
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
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}
	}
}
