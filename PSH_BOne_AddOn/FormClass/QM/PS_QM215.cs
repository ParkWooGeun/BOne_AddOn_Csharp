using System;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 사내표준관리대장
	/// </summary>
	internal class PS_QM215 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;

		private SAPbouiCOM.DBDataSource oDS_PS_QM215L; //라인(개정내역)
		private SAPbouiCOM.DBDataSource oDS_PS_QM215M; //라인(세부내역)

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM215.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM215_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM215");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_QM215_CreateItems();
				PS_QM215_ComboBox_Setting();
				PS_QM215_SetDocument(oFormDocEntry);
				PS_QM215_FormResize();
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
		/// PS_QM215_CreateItems
		/// </summary>
		private void PS_QM215_CreateItems()
		{
			try
			{
				oDS_PS_QM215L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oDS_PS_QM215M = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");

				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				oMat02 = oForm.Items.Item("Mat02").Specific;
				oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat02.AutoResizeColumns();

				//사업부
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				//부서
				oForm.DataSources.UserDataSources.Add("DeptCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("DeptCode").Specific.DataBind.SetBound(true, "", "DeptCode");

				//부서명
				oForm.DataSources.UserDataSources.Add("DeptName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("DeptName").Specific.DataBind.SetBound(true, "", "DeptName");

				//폐기문서 제외
				oForm.DataSources.UserDataSources.Add("chkDelYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("chkDelYN").Specific.DataBind.SetBound(true, "", "chkDelYN");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM215_ComboBox_Setting
		/// </summary>
		private void PS_QM215_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업부
				oForm.Items.Item("BPLId").Specific.ValidValues.Add("", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", false, false);
				oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM215_SetDocument
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		private void PS_QM215_SetDocument(string oFromDocEntry01)
		{
			try
			{
				if (string.IsNullOrEmpty(oFromDocEntry01))
				{
					PS_QM215_FormItemEnabled();
				}
				else
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM215_FormResize
		/// </summary>
		private void PS_QM215_FormResize()
		{
			try
			{
				oForm.Items.Item("Rec01").Width = oForm.Items.Item("Mat01").Width + 20;
				oForm.Items.Item("Rec01").Height = oForm.Items.Item("Mat01").Height + 50;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM215_FormItemEnabled
		/// </summary>
		private void PS_QM215_FormItemEnabled()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
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
		/// PS_QM215_AddMatrixRow02
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_QM215_AddMatrixRow02(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oDS_PS_QM215M.InsertRecord(oRow);
				}
				oMat02.AddRow();
				oDS_PS_QM215M.Offset = oRow;
				oMat02.LoadFromDataSource();
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
		/// PS_QM215_MTX01
		/// </summary>
		private void PS_QM215_MTX01()
		{
			int loopCount;

			int BPLId;       //사업부
			string DeptCode; //부서
			string chkDelYN; //폐기문서제외
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				BPLId = Convert.ToInt32(oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim());
				DeptCode = oForm.Items.Item("DeptCode").Specific.Value.ToString().Trim();
				if (oForm.Items.Item("chkDelYN").Specific.Checked == true)
				{
					chkDelYN = "1";
				}
				else
				{
					chkDelYN = "0";
				}

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC PS_QM215_01 '" + BPLId + "','" + DeptCode + "','" + chkDelYN + "'";
				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oMat01.Clear();
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount != 0)
					{
						oDS_PS_QM215L.InsertRecord(loopCount);
					}
					oDS_PS_QM215L.Offset = loopCount;
					oDS_PS_QM215L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));                             //라인번호
					oDS_PS_QM215L.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("Level").Value.ToString().Trim());    //Level
					oDS_PS_QM215L.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("StdName").Value.ToString().Trim());  //표준명
					oDS_PS_QM215L.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("DocCode").Value.ToString().Trim());  //문서코드
					oDS_PS_QM215L.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("DeptCode").Value.ToString().Trim()); //부서코드
					oDS_PS_QM215L.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("DeptName").Value.ToString().Trim()); //원안작성부서
					oDS_PS_QM215L.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("StrDate").Value.ToString().Trim());  //제정일자
					oDS_PS_QM215L.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("MaxDate").Value.ToString().Trim());  //최종개정일자
					oDS_PS_QM215L.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("DelDate").Value.ToString().Trim());  //폐기일자
					oDS_PS_QM215L.SetValue("U_ColReg09", loopCount, oRecordSet.Fields.Item("CrtDim").Value.ToString().Trim());   //개정차수
					oRecordSet.MoveNext();

					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();
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
		/// PS_QM215_MTX02
		/// </summary>
		private void PS_QM215_MTX02()
		{
			int loopCount;

			int BPLId;       //사업부
			string DeptCode; //부서
			string chkDelYN; //폐기문서제외
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				BPLId = Convert.ToInt32(oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim());
				DeptCode = oForm.Items.Item("DeptCode").Specific.Value.ToString().Trim();
				if (oForm.Items.Item("chkDelYN").Specific.Checked == true)
				{
					chkDelYN = "1";
				}
				else
				{
					chkDelYN = "0";
				}

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC PS_QM215_02 '" + BPLId + "','" + DeptCode + "','" + chkDelYN + "'";
				oRecordSet.DoQuery(sQry);

				oMat02.Clear();
				oMat02.FlushToDataSource();
				oMat02.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oMat02.Clear();
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount != 0)
					{
						oDS_PS_QM215M.InsertRecord(loopCount);
					}
					oDS_PS_QM215M.Offset = loopCount;
					oDS_PS_QM215M.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));                             //라인번호
					oDS_PS_QM215M.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("Level").Value.ToString().Trim());    //Level
					oDS_PS_QM215M.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("StdName").Value.ToString().Trim());  //표준명
					oDS_PS_QM215M.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("DocCode").Value.ToString().Trim());  //문서코드
					oDS_PS_QM215M.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("DeptCode").Value.ToString().Trim()); //부서코드
					oDS_PS_QM215M.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("DeptName").Value.ToString().Trim()); //원안작성부서
					oDS_PS_QM215M.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("Cmt").Value.ToString().Trim());      //내역
					oDS_PS_QM215M.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("Date").Value.ToString().Trim());     //일자
					oDS_PS_QM215M.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("CrtCmt").Value.ToString().Trim());   //개정내용
					oRecordSet.MoveNext();

					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat02.LoadFromDataSource();
				oMat02.AutoResizeColumns();
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
		/// PS_QM215_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_QM215_Print_Report01()
		{
			string WinTitle;
			string ReportName;
			int BPLId;       //사업부
			string DeptCode; //부서
			string chkDelYN; //폐기문서제외
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLId = Convert.ToInt32(oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim());
				DeptCode = oForm.Items.Item("DeptCode").Specific.Value.ToString().Trim();
				if (oForm.Items.Item("chkDelYN").Specific.Checked == true)
				{
					chkDelYN = "1";
				}
				else
				{
					chkDelYN = "0";
				}

				WinTitle = "[PS_QM215] 레포트(개정내역)";
				ReportName = "PS_QM215_01.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
				dataPackParameter.Add(new PSH_DataPackClass("@DeptCode", DeptCode));
				dataPackParameter.Add(new PSH_DataPackClass("@chkDelYN", chkDelYN));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM215_Print_Report02
		/// </summary>
		[STAThread]
		private void PS_QM215_Print_Report02()
		{
			string WinTitle;
			string ReportName;
			int BPLId;       //사업부
			string DeptCode; //부서
			string chkDelYN; //폐기문서제외
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLId = Convert.ToInt32(oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim());
				DeptCode = oForm.Items.Item("DeptCode").Specific.Value.ToString().Trim();
				if (oForm.Items.Item("chkDelYN").Specific.Checked == true)
				{
					chkDelYN = "1";
				}
				else
				{
					chkDelYN = "0";
				}

				WinTitle = "[PS_QM215] 레포트(세부내역)";
				ReportName = "PS_QM215_02.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
				dataPackParameter.Add(new PSH_DataPackClass("@DeptCode", DeptCode));
				dataPackParameter.Add(new PSH_DataPackClass("@chkDelYN", chkDelYN));

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
				//	Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
				//    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
				//	Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_CLICK: //6
				//    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
					Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8	
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
					Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
					Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
					break;
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
					if (pVal.ItemUID == "btnSearch")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_QM215_MTX01(); //매트릭스에 데이터 로드(개정내역)
							PS_QM215_MTX02(); //매트릭스에 데이터 로드(세부내역)
						}
					}
					else if (pVal.ItemUID == "btnPrint01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_QM215_Print_Report01);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
					else if (pVal.ItemUID == "btnPrint02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_QM215_Print_Report02);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					//폴더를 사용할 때는 필수 소스
					if (pVal.ItemUID == "Folder01")
					{
						oForm.PaneLevel = 1;
					}
					if (pVal.ItemUID == "Folder02")
					{
						oForm.PaneLevel = 2;
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "DeptCode", "");
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_EVENT_MATRIX_LINK_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			//string FileName;

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.ColUID == "DocCode")
						{
							//사용안함 (Convetion 하지않음)
							//FileName = "\\\\191.1.1.220\\WorkSpace\\사내규정\\" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + ".pdf");
							//MessageAPIs.ShellExecute(0, "Open", "AcroRd32.exe", FileName, "", 5);
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
		/// VALIDATE 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "DeptCode")
						{
							sQry = " SELECT      T1.U_CodeNm AS [Name]";
							sQry += "  FROM        [@PS_HR200H] AS T0";
							sQry += "                  INNER JOIN";
							sQry += "                  [@PS_HR200L] AS T1";
							sQry += "                      ON T0.Code = T1.Code";
							sQry += "  WHERE       T1.U_UseYN = 'Y'";
							sQry += "                  AND T0.Name = '부서'";
							sQry += "                  AND T1.U_Code = '" + oForm.Items.Item("DeptCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("DeptName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
					PS_QM215_FormItemEnabled();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
					PS_QM215_FormResize();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM215L);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM215M);
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
						case "1283": //삭제
							break;
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
							PS_QM215_AddMatrixRow02(oMat01.VisualRowCount, false);
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
							oDS_PS_QM215L.RemoveRecord(oDS_PS_QM215L.Size - 1);
							oMat01.LoadFromDataSource();
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
	}
}
