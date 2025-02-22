using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 구매현황조회
	/// </summary>
	internal class PS_MM965 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM965.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM965_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM965");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_MM965_CreateItems();
				PS_MM965_ComboBox_Setting();
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
		/// PS_MM965_CreateItems
		/// </summary>
		private void PS_MM965_CreateItems()
		{
			try
			{
				oGrid = oForm.Items.Item("Grid01").Specific;

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				//기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");

				//기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");

				//품의구분
				oForm.DataSources.UserDataSources.Add("Purchase", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("Purchase").Specific.DataBind.SetBound(true, "", "Purchase");

				//품의완료여부
				oForm.DataSources.UserDataSources.Add("OrderYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("OrderYN").Specific.DataBind.SetBound(true, "", "OrderYN");

				//검수완료여부
				oForm.DataSources.UserDataSources.Add("InputYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("InputYN").Specific.DataBind.SetBound(true, "", "InputYN");

				//구분
				oForm.DataSources.UserDataSources.Add("Class", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("Class").Specific.DataBind.SetBound(true, "", "Class");

				oForm.Items.Item("FrDt").Specific.Value = DateTime.Now.ToString("yyyyMM01");
				oForm.Items.Item("ToDt").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM965_ComboBox_Setting
		/// </summary>
		private void PS_MM965_ComboBox_Setting()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//품의구분
				sQry = "    SELECT       Code AS [Code],";
				sQry += "                 Name AS [Name]";
				sQry += " FROM        [@PSH_ORDTYP]";
				sQry += " WHERE       Code IN ('10','20','30','40')";
				sQry += " ORDER BY  Code";
				dataHelpClass.Set_ComboList(oForm.Items.Item("Purchase").Specific, sQry, "", false, false);
				oForm.Items.Item("Purchase").Specific.Select("10", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//품의완료여부
				oForm.Items.Item("OrderYN").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("OrderYN").Specific.ValidValues.Add("Y", "품의완료");
				oForm.Items.Item("OrderYN").Specific.ValidValues.Add("N", "품의미완료");
				oForm.Items.Item("OrderYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//검수완료여부
				oForm.Items.Item("InputYN").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("InputYN").Specific.ValidValues.Add("Y", "검수완료");
				oForm.Items.Item("InputYN").Specific.ValidValues.Add("N", "검수미완료");
				oForm.Items.Item("InputYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//구분
				oForm.Items.Item("Class").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("Class").Specific.ValidValues.Add("1", "공구");
				oForm.Items.Item("Class").Specific.ValidValues.Add("2", "장비");
				oForm.Items.Item("Class").Specific.ValidValues.Add("3", "몰드");
				oForm.Items.Item("Class").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 메트릭스에 데이터 로드
		/// </summary>
		private void PS_MM965_MTX01()
		{
			string BPLID;
			string FrDt;
			string ToDt;
			string Purchase;
			string OrderYN;
			string InputYN;
			string Class;
			string errMessage = string.Empty;
			string sQry = string.Empty;
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				Purchase = oForm.Items.Item("Purchase").Specific.Value.ToString().Trim();
				OrderYN = oForm.Items.Item("OrderYN").Specific.Value.ToString().Trim();
				InputYN = oForm.Items.Item("InputYN").Specific.Value.ToString().Trim();
				Class = oForm.Items.Item("Class").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				if (Purchase == "10") //원자재품의
				{
					sQry = "       EXEC PS_MM965_10 '";
					sQry += BPLID + "','";
					sQry += FrDt + "','";
					sQry += ToDt + "','";
					sQry += Purchase + "','";
					sQry += OrderYN + "','";
					sQry += InputYN + "','";
					sQry += Class + "','";
					sQry += 1 + "'"; //조회는 1
				}
				else if (Purchase == "20") //부자재품의
				{
					sQry = "       EXEC PS_MM965_20 '";
					sQry += BPLID + "','";
					sQry += FrDt + "','";
					sQry += ToDt + "','";
					sQry += Purchase + "','";
					sQry += OrderYN + "','";
					sQry += InputYN + "','";
					sQry += Class + "','";
					sQry += 1 + "'"; //조회는 1
				}
				else if (Purchase == "30") //가공비품의
				{
					sQry = "       EXEC PS_MM965_30 '";
					sQry += BPLID + "','";
					sQry += FrDt + "','";
					sQry += ToDt + "','";
					sQry += Purchase + "','";
					sQry += Class + "','";
					sQry += 1 + "'"; //조회는 1
				}
				else if (Purchase == "40") //외주제작품의
				{
					sQry = "       EXEC PS_MM965_40 '";
					sQry += BPLID + "','";
					sQry += FrDt + "','";
					sQry += ToDt + "','";
					sQry += Purchase + "','";
					sQry += Class + "','";
					sQry += 1 + "'"; //조회는 1
				}

				oGrid.DataTable.Clear();
				oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry);
				oGrid.DataTable = oForm.DataSources.DataTables.Item("DataTable");
				
				if (Purchase == "10") //원자재품의
				{
					oGrid.Columns.Item(8).RightJustified = true;
					oGrid.Columns.Item(14).RightJustified = true;
					oGrid.Columns.Item(17).RightJustified = true;
					oGrid.Columns.Item(18).RightJustified = true;
					oGrid.Columns.Item(20).RightJustified = true;
					oGrid.Columns.Item(22).RightJustified = true;
					oGrid.Columns.Item(23).RightJustified = true;

					oGrid.Columns.Item(9).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125));  //수주일, 노랑
					oGrid.Columns.Item(10).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //요청일, 노랑
					oGrid.Columns.Item(13).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //견적일, 노랑
					oGrid.Columns.Item(15).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 210, 255));   //차이(요청-견적), 하늘
					oGrid.Columns.Item(16).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //품의일, 노랑
					oGrid.Columns.Item(19).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 210, 255));   //차이(견적-품의), 하늘
					oGrid.Columns.Item(20).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //가입고일, 노랑
					oGrid.Columns.Item(21).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 210, 255));   //차이(품의-가입고), 하늘
					oGrid.Columns.Item(22).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //검수입고일, 노랑
					oGrid.Columns.Item(23).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 210, 255));   //차이(가입고-검수), 하늘
					oGrid.Columns.Item(24).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 167, 167)); //총소요일, 빨강
				}
				else if (Purchase == "20") //부자재품의
				{
					oGrid.Columns.Item(3).RightJustified = true;
					oGrid.Columns.Item(8).RightJustified = true;
					oGrid.Columns.Item(11).RightJustified = true;
					oGrid.Columns.Item(12).RightJustified = true;
					oGrid.Columns.Item(14).RightJustified = true;
					oGrid.Columns.Item(16).RightJustified = true;
					oGrid.Columns.Item(17).RightJustified = true;

					oGrid.Columns.Item(4).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125));  //요청일, 노랑
					oGrid.Columns.Item(7).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125));  //견적일, 노랑
					oGrid.Columns.Item(9).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 210, 255));	   //차이(요청-견적), 하늘
					oGrid.Columns.Item(10).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125));  //품의일, 노랑
					oGrid.Columns.Item(13).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 210, 255));   //차이(견적-품의), 하늘
					oGrid.Columns.Item(14).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //가입고일, 노랑
					oGrid.Columns.Item(15).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 210, 255));   //차이(품의-가입고), 하늘
					oGrid.Columns.Item(16).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //검수입고일, 노랑
					oGrid.Columns.Item(17).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 210, 255));   //차이(가입고-검수), 하늘
					oGrid.Columns.Item(18).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 167, 167)); //총소요일, 빨강
				}
				else if (Purchase == "30") //가공비품의
				{
					oGrid.Columns.Item(6).RightJustified = true;
					oGrid.Columns.Item(7).RightJustified = true;
					oGrid.Columns.Item(12).RightJustified = true;
					oGrid.Columns.Item(14).RightJustified = true;
					oGrid.Columns.Item(15).RightJustified = true;

					oGrid.Columns.Item(9).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125));  //수주일, 노랑
					oGrid.Columns.Item(12).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //품의일, 노랑
					oGrid.Columns.Item(13).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //가입고일, 노랑
					oGrid.Columns.Item(14).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 210, 255));   //차이(품의-가입고), 하늘
					oGrid.Columns.Item(15).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //검수입고일, 노랑
					oGrid.Columns.Item(16).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 210, 255));   //차이(가입고-품의), 하늘
					oGrid.Columns.Item(17).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 167, 167)); //총소요일, 빨강
				}
				else if (Purchase == "40") //외주제작품의
				{
					oGrid.Columns.Item(5).RightJustified = true;
					oGrid.Columns.Item(6).RightJustified = true;
					oGrid.Columns.Item(11).RightJustified = true;
					oGrid.Columns.Item(13).RightJustified = true;
					oGrid.Columns.Item(14).RightJustified = true;

					oGrid.Columns.Item(8).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125));  //수주일, 노랑
					oGrid.Columns.Item(11).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125));  //품의일, 노랑
					oGrid.Columns.Item(12).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //가입고일, 노랑
					oGrid.Columns.Item(13).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 210, 255));   //차이(품의-가입고), 하늘
					oGrid.Columns.Item(14).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //검수입고일, 노랑
					oGrid.Columns.Item(15).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 210, 255));   //차이(가입고-품의), 하늘
					oGrid.Columns.Item(16).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 167, 167)); //총소요일, 빨강
				}

				if (oGrid.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}
				oGrid.AutoResizeColumns();
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
		/// PS_MM965_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_MM965_Print_Report01()
		{
			string WinTitle;
			string ReportName = string.Empty;
			string BPLID;
			string FrDt;
			string ToDt;
			string Purchase;
			string OrderYN;
			string InputYN;
			string Class;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				Purchase = oForm.Items.Item("Purchase").Specific.Value.ToString().Trim();
				OrderYN = oForm.Items.Item("OrderYN").Specific.Value.ToString().Trim();
				InputYN = oForm.Items.Item("InputYN").Specific.Value.ToString().Trim();
				Class = oForm.Items.Item("Class").Specific.Value.ToString().Trim();

				WinTitle = "[PS_MM965] 레포트";
				
				if (Purchase == "10") //원재료품의
				{
					ReportName = "PS_MM965_10.rpt";
				}
				else if (Purchase == "20") //부재료품의
				{
					ReportName = "PS_MM965_20.rpt";
				}
				else if (Purchase == "30") //가공비품의
				{
					ReportName = "PS_MM965_30.rpt";
				}
				else if (Purchase == "40") //'외주제작품의
				{
					ReportName = "PS_MM965_40.rpt";
				}

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
				//Parameter
				if (Purchase == "10" || Purchase == "20")
				{
					dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
					dataPackParameter.Add(new PSH_DataPackClass("@FrDt", DateTime.ParseExact(FrDt, "yyyyMMdd", null)));
					dataPackParameter.Add(new PSH_DataPackClass("@ToDt", DateTime.ParseExact(ToDt, "yyyyMMdd", null)));
					dataPackParameter.Add(new PSH_DataPackClass("@Purchase", Purchase));
					dataPackParameter.Add(new PSH_DataPackClass("@OrderYN", OrderYN));
					dataPackParameter.Add(new PSH_DataPackClass("@InputYN", InputYN));
					dataPackParameter.Add(new PSH_DataPackClass("@Class", Class));
					dataPackParameter.Add(new PSH_DataPackClass("@Mode", 2));
				}
				else if (Purchase == "30" || Purchase == "40")
				{
					dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
					dataPackParameter.Add(new PSH_DataPackClass("@FrDt", DateTime.ParseExact(FrDt, "yyyyMMdd", null)));
					dataPackParameter.Add(new PSH_DataPackClass("@ToDt", DateTime.ParseExact(ToDt, "yyyyMMdd", null)));
					dataPackParameter.Add(new PSH_DataPackClass("@Purchase", Purchase));
					dataPackParameter.Add(new PSH_DataPackClass("@Class", Class));
					dataPackParameter.Add(new PSH_DataPackClass("@Mode", 2));
				}
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
				//case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
				//	Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
				//	Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
				//    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
				//	Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_CLICK: //6
				//	Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
				//    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
				//    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
				//	Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
				//case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
				//    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
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
					if (pVal.ItemUID == "BtnSearch")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_MM965_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnPrint")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_MM965_Print_Report01);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid);
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
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1283": //삭제
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1293": //행삭제
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1287": // 복제
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1293": //행삭제
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
