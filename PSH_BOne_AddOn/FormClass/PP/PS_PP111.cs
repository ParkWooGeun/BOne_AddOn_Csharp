using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 추가예상공수(비용)등록대상
	/// </summary>
	internal class PS_PP111 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP111L;  //등록라인

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP111.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP111_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP111");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

				oForm.Freeze(true);

				PS_PP111_CreateItems();
				PS_PP111_SetComboBox();
				PS_PP111_Initialize();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", false); // 행삭제
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
		/// PS_PP111_Initialize
		/// </summary>
		private void PS_PP111_Initialize()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); // 아이디별 사업장 세팅
				oForm.Items.Item("DocYM").Specific.Value = DateTime.Now.ToString("yyyyMM"); // 기본년월
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP111_CreateItems
		/// </summary>
		private void PS_PP111_CreateItems()
		{
			try
			{
				oDS_PS_PP111L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

				// 메트릭스 개체 할당
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				//기준년월
				oForm.DataSources.UserDataSources.Add("DocYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("DocYM").Specific.DataBind.SetBound(true, "", "DocYM");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP111_SetComboBox
		/// </summary>
		private void PS_PP111_SetComboBox()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
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
		/// PS_PP111_MTX01
		/// </summary>
		private void PS_PP111_MTX01()
		{
			short i;
			string sQry;
			string errMessage = string.Empty;

			string BPLID; //사업장
			string StdYM; //기준년월

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				StdYM = oForm.Items.Item("DocYM").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = "EXEC [PS_PP111_01] '";
				sQry += BPLID + "','";
				sQry += StdYM + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_PP111L.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP111L.Size)
					{
						oDS_PS_PP111L.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_PP111L.Offset = i;

					oDS_PS_PP111L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP111L.SetValue("U_ColReg11", i, "N"); //선택
					oDS_PS_PP111L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("PoEntry").Value.ToString().Trim());     //작지문서번호
					oDS_PS_PP111L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("PoLine").Value.ToString().Trim());      //공정순번
					oDS_PS_PP111L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());      //작번
					oDS_PS_PP111L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("OrdSub1").Value.ToString().Trim());     //서브작번1
					oDS_PS_PP111L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("OrdSub2").Value.ToString().Trim());     //서브작번2
					oDS_PS_PP111L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());    //품목명
					oDS_PS_PP111L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());      //공정코드
					oDS_PS_PP111L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());      //공정명
					oDS_PS_PP111L.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("InVal").Value.ToString().Trim());       //발생(비용/공수)
					oDS_PS_PP111L.SetValue("U_ColQty02", i, oRecordSet.Fields.Item("ReVal").Value.ToString().Trim());       //추가발생[외주제작](비용/공수)
					oDS_PS_PP111L.SetValue("U_ColQty03", i, oRecordSet.Fields.Item("ReVal2").Value.ToString().Trim());      //추가발생[외주가공](비용/공수)
					oDS_PS_PP111L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("CreateUser").Value.ToString().Trim());  //등록자
					oDS_PS_PP111L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("UpdateUser").Value.ToString().Trim());  //수정자
					oDS_PS_PP111L.SetValue("U_ColReg11", i, oRecordSet.Fields.Item("CUName").Value.ToString().Trim());      //등록자
					oDS_PS_PP111L.SetValue("U_ColReg12", i, oRecordSet.Fields.Item("UUName").Value.ToString().Trim());      //수정자

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";

					oMat.LoadFromDataSource();
					oMat.AutoResizeColumns();
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
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP111_AddData
		/// </summary>
		private void PS_PP111_AddData()
		{
			short loopCount;
			string sQry;
			string BPLID;		//사업장
			string StdYM;		//기준년월
			string POEntry;		//작지문서번호
			string POLine;		//공정순번
			string OrdNum;		//작번
			string OrdSub1;		//서브작번1
			string OrdSub2;		//서브작번2
			string ItemName;	//품목명
			string CpCode;		//공정코드
			string CpName;		//공정명
			decimal InVal;		//발생(비용/공수)
			decimal ReVal;		//추가발생[외주제작](비용/공수)
			decimal ReVal2;		//추가발생[외주가공]
			string UserSign;    //UserSign

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				StdYM = oForm.Items.Item("DocYM").Specific.Value.ToString().Trim();
				UserSign = PSH_Globals.oCompany.UserSignature.ToString();

				oMat.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					if (oMat.Columns.Item("Check").Cells.Item(loopCount + 1).Specific.Checked == true)
					{
						POEntry = oDS_PS_PP111L.GetValue("U_ColReg01", loopCount).ToString().Trim();                    //작지문서번호
						POLine = oDS_PS_PP111L.GetValue("U_ColReg02", loopCount).ToString().Trim();                     //공정순번
						OrdNum = oDS_PS_PP111L.GetValue("U_ColReg03", loopCount).ToString().Trim();                     //작번
						OrdSub1 = oDS_PS_PP111L.GetValue("U_ColReg04", loopCount).ToString().Trim();                    //서브작번1
						OrdSub2 = oDS_PS_PP111L.GetValue("U_ColReg05", loopCount).ToString().Trim();                    //서브작번2
						ItemName = oDS_PS_PP111L.GetValue("U_ColReg06", loopCount).ToString().Trim();                   //품목명
						CpCode = oDS_PS_PP111L.GetValue("U_ColReg07", loopCount).ToString().Trim();                     //공정코드
						CpName = oDS_PS_PP111L.GetValue("U_ColReg08", loopCount).ToString().Trim();                     //공정명
						InVal = Convert.ToDecimal(oDS_PS_PP111L.GetValue("U_ColQty01", loopCount).ToString().Trim());   //발생(비용/공수)
						ReVal = Convert.ToDecimal(oDS_PS_PP111L.GetValue("U_ColQty02", loopCount).ToString().Trim());   //추가발생[외주제작](비용/공수)
						ReVal2 = Convert.ToDecimal(oDS_PS_PP111L.GetValue("U_ColQty03", loopCount).ToString().Trim());  //추가발생[외주가공]

						ProgressBar01.Text = "저장 중...";

						sQry = " EXEC [PS_PP111_02] ";
						sQry += "'" + BPLID + "',";     //사업장
						sQry += "'" + StdYM + "',";     //기준년월
						sQry += "'" + POEntry + "',";   //작지문서번호
						sQry += "'" + POLine + "',";    //공정순번
						sQry += "'" + OrdNum + "',";    //작번
						sQry += "'" + OrdSub1 + "',";   //서브작번1
						sQry += "'" + OrdSub2 + "',";   //서브작번2
						sQry += "'" + ItemName + "',";  //품목명
						sQry += "'" + CpCode + "',";    //공정코드
						sQry += "'" + CpName + "',";    //공정명
						sQry += "'" + InVal + "',";     //발생(비용/공수)
						sQry += "'" + ReVal + "',";     //추가발생[외주제작](비용/공수)
						sQry += "'" + ReVal2 + "',";    //추가발생[외주가공](비용/공수)
						sQry += "'" + UserSign + "'";   //UserSign

						oRecordSet.DoQuery(sQry);
					}
				}
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
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP111_PrintReport
		/// </summary>
		[STAThread]
		private void PS_PP111_PrintReport()
		{
			string sQry;
			string BPLName;
			string WinTitle;
			string ReportName;

			string BPLID;
			string DocYM;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocYM = oForm.Items.Item("DocYM").Specific.Value.ToString().Trim();

				sQry = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" + BPLID + "'";
				oRecordSet.DoQuery(sQry);
				BPLName = oRecordSet.Fields.Item(0).Value.ToString().Trim();

				WinTitle = "추가예상공수(비용)등록대상[PS_PP111_01]";
				ReportName = "PS_PP111_01.RPT";

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드
				dataPackFormula.Add(new PSH_DataPackClass("@BPLId", BPLID));
				dataPackFormula.Add(new PSH_DataPackClass("@StdYM", DocYM));

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@StdYM", DocYM));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "1")
					{
					}
					else if (pVal.ItemUID == "BtnSrch")
					{
						PS_PP111_MTX01();
					}
					else if (pVal.ItemUID == "BtnSave")
					{
						PS_PP111_AddData();
					}
					else if (pVal.ItemUID == "BtnPrint")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP111_PrintReport);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP111L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
