using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 제품재고수불등록(년)
	/// </summary>
	internal class PS_MM200 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_MM200H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_MM200L; //등록라인
		private int oLastColRow01;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM200.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}
				oFormUniqueID = "PS_MM200_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM200");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "Code"; //UDO방식일때
				
				oForm.Freeze(true);

				PS_MM200_CreateItems();
				PS_MM200_ComboBox_Setting();
				PS_MM200_SetDocument(oFormDocEntry);

				oForm.EnableMenu("1293", true); // 행삭제
				oForm.EnableMenu("1287", true); // 복제
				oForm.EnableMenu("1284", true); // 취소
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
		/// PS_MM200_CreateItems
		/// </summary>
		private void PS_MM200_CreateItems()
		{
			try
			{
				oDS_PS_MM200H = oForm.DataSources.DBDataSources.Item("@PS_MM200H");
				oDS_PS_MM200L = oForm.DataSources.DBDataSources.Item("@PS_MM200L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM200_ComboBox_Setting
		/// </summary>
		private void PS_MM200_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				//아이디별 사업장 세팅
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				oForm.Items.Item("Gubun").Specific.ValidValues.Add("10", "제품");
				oForm.Items.Item("Gubun").Specific.ValidValues.Add("20", "제품(임가공)");
				oForm.Items.Item("Gubun").Specific.ValidValues.Add("30", "상품");
				oForm.Items.Item("Gubun").Specific.ValidValues.Add("40", "원재료");
				oDS_PS_MM200H.SetValue("U_Gubun", 0, "10");
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
		/// PS_MM200_SetDocument
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		private void PS_MM200_SetDocument(string oFromDocEntry01)
		{
			try
			{
				if (string.IsNullOrEmpty(oFromDocEntry01))
				{
					PS_MM200_FormItemEnabled();
					PS_MM200_AddMatrixRow(0, true);
				}
				else
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					PS_MM200_FormItemEnabled();
					oForm.Items.Item("Code").Specific.Value = oFromDocEntry01;
					oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM200_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_MM200_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				//행추가여부
				if (RowIserted == false)
				{
					oRow = oMat.RowCount;
					oDS_PS_MM200L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_MM200L.Offset = oRow;
				oDS_PS_MM200L.SetValue("LineId", oRow, Convert.ToString(oRow + 1));
				oDS_PS_MM200L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		/// PS_MM200_FormItemEnabled
		/// </summary>
		private void PS_MM200_FormItemEnabled()
		{
			try
			{
				oForm.Freeze(true);
				//각모드에따른 아이템설정
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.EnableMenu("1281", true);	 //찾기
					oForm.EnableMenu("1282", false); //추가
					oForm.Items.Item("Code").Enabled = false;
					oForm.Items.Item("YEAR").Enabled = true;
					oForm.Items.Item("Gubun").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.EnableMenu("1281", true); //찾기
					oForm.Items.Item("Code").Enabled = false;
					oForm.Items.Item("YEAR").Enabled = true;
					oForm.Items.Item("Gubun").Enabled = true;
					oForm.EnableMenu("1282", true); //추가
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("Code").Enabled = false;
					oForm.Items.Item("YEAR").Enabled = false;
					oForm.Items.Item("Gubun").Enabled = false;
					oForm.EnableMenu("1282", true);//추가
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
		/// PS_MM200_MatrixSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_MM200_MatrixSpaceLineDel()
		{
			bool functionReturnValue = false;
			int i;
			string errMessage = string.Empty;

			try
			{
				oMat.FlushToDataSource();

				if (oMat.VisualRowCount == 0)
				{
					errMessage = "라인데이타가 없습니다. 확인하세요.";
					throw new Exception();
				}
				else if (oMat.VisualRowCount == 1)
				{
					if (string.IsNullOrEmpty(oDS_PS_MM200L.GetValue("U_ItemCode", 0).ToString().Trim()))
					{
						errMessage = "라인데이타가 없습니다. 확인하세요.";
						throw new Exception();
					}
				}

				if (oMat.VisualRowCount > 0)
				{
					for (i = 0; i <= oMat.VisualRowCount - 2; i++)
					{
						oDS_PS_MM200L.Offset = i;
						if (string.IsNullOrEmpty(oDS_PS_MM200L.GetValue("U_ItemCode", i).ToString().Trim()))
						{
							errMessage = "품목코드는 필수입력사항입니다. 확인하세요.";
							throw new Exception();
						}
					}

					if (string.IsNullOrEmpty(oDS_PS_MM200L.GetValue("U_ItemCode", oMat.VisualRowCount - 1).ToString().Trim()))
					{
						oDS_PS_MM200L.RemoveRecord(oMat.VisualRowCount - 1);
					}
				}

				oMat.LoadFromDataSource();
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
		/// PS_MM200_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_MM200_HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_MM200H.GetValue("U_YEAR", 0).ToString().Trim()))
				{
					errMessage = "년도는 필수입력 사항입니다.";
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
			return functionReturnValue;
		}

		/// <summary>
		/// PS_MM200_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_MM200_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				switch (oUID)
				{
					case "":
						break;

					case "Mat01":
						if (oCol == "ItemCode")
						{
							if ((oRow == oMat.RowCount || oMat.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim()))
							{
								oMat.FlushToDataSource();
								PS_MM200_AddMatrixRow(oMat.RowCount, false);
								oMat.Columns.Item("ItemCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							}

							sQry = "Select FrgnName, U_Size, InvntryUom From OITM Where ItemCode = '" + oMat.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim() +"'";
							oRecordSet.DoQuery(sQry);
							oMat.Columns.Item("FrgnName").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
							oMat.Columns.Item("Size").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
							oMat.Columns.Item("Unit").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item(2).Value.ToString().Trim();
							oMat.FlushToDataSource();
						}
						break;
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
		/// PS_MM200_Excel_Upload
		/// </summary>
		[STAThread]
		private void PS_MM200_Excel_Upload()
		{
			int rowCount;
			int loopCount;
			string sFile;
			bool sucessFlag = false;
			short columnCount = 13; //엑셀 컬럼수
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			CommonOpenFileDialog commonOpenFileDialog = new CommonOpenFileDialog();

			commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("Excel Files", "*.xls;*.xlsx"));
			commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("모든 파일", "*.*"));

			if (commonOpenFileDialog.ShowDialog() == CommonFileDialogResult.Ok)
			{
				sFile = commonOpenFileDialog.FileName;
			}
			else //Cancel 버튼 클릭
			{
				return;
			}

			if (string.IsNullOrEmpty(sFile))
			{
				return;
			}

			//엑셀 Object 연결
			//암시적 객체참조 시 Excel.exe 메모리 반환이 안됨, 아래와 같이 명시적 참조로 선언
			Microsoft.Office.Interop.Excel.ApplicationClass xlapp = new Microsoft.Office.Interop.Excel.ApplicationClass();
			Microsoft.Office.Interop.Excel.Workbooks xlwbs = xlapp.Workbooks;
			Microsoft.Office.Interop.Excel.Workbook xlwb = xlwbs.Open(sFile);
			Microsoft.Office.Interop.Excel.Sheets xlshs = xlwb.Worksheets;
			Microsoft.Office.Interop.Excel.Worksheet xlsh = (Microsoft.Office.Interop.Excel.Worksheet)xlshs[1];
			Microsoft.Office.Interop.Excel.Range xlCell = xlsh.Cells;
			Microsoft.Office.Interop.Excel.Range xlRange = xlsh.UsedRange;
			Microsoft.Office.Interop.Excel.Range xlRow = xlRange.Rows;

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("시작!", xlRow.Count, false);

			oForm.Freeze(true);

			oMat.Clear();
			oMat.FlushToDataSource();
			oMat.LoadFromDataSource();

			try
			{
				for (rowCount = 2; rowCount <= xlRow.Count; rowCount++)
				{
					if (rowCount - 2 != 0)
					{
						oDS_PS_MM200L.InsertRecord(rowCount - 2);
					}

					Microsoft.Office.Interop.Excel.Range[] r = new Microsoft.Office.Interop.Excel.Range[columnCount + 1];

					for (loopCount = 1; loopCount <= columnCount; loopCount++)
					{
						r[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[rowCount, loopCount];
					}

					sQry = "Select FrgnName, U_Size, InvntryUom From OITM Where ItemCode = '" + Convert.ToString(r[1].Value) + "'";
					oRecordSet.DoQuery(sQry);

					oDS_PS_MM200L.Offset = rowCount - 2;
					oDS_PS_MM200L.SetValue("U_LineNum", rowCount - 2, Convert.ToString(rowCount - 1));
					oDS_PS_MM200L.SetValue("U_ItemCode", rowCount - 2, Convert.ToString(r[1].Value)); //코드
					oDS_PS_MM200L.SetValue("U_FrgnName", rowCount - 2, oRecordSet.Fields.Item(0).Value.ToString().Trim()); //품명
					oDS_PS_MM200L.SetValue("U_Size", rowCount - 2, oRecordSet.Fields.Item(1).Value.ToString().Trim()); //규격
					oDS_PS_MM200L.SetValue("U_Unit", rowCount - 2, oRecordSet.Fields.Item(2).Value.ToString().Trim()); //단위
					oDS_PS_MM200L.SetValue("U_iwqty", rowCount - 2, Convert.ToString(r[2].Value));    //기초수량
					oDS_PS_MM200L.SetValue("U_iwamt", rowCount - 2, Convert.ToString(r[3].Value));    //기초금액
					oDS_PS_MM200L.SetValue("U_i1qty", rowCount - 2, Convert.ToString(r[4].Value));    //입고수량
					oDS_PS_MM200L.SetValue("U_i1amt", rowCount - 2, Convert.ToString(r[5].Value));    //입고금액
					oDS_PS_MM200L.SetValue("U_i2qty", rowCount - 2, Convert.ToString(r[6].Value));    //타계정입고수량
					oDS_PS_MM200L.SetValue("U_i2amt", rowCount - 2, Convert.ToString(r[7].Value));    //타계정입고금액
					oDS_PS_MM200L.SetValue("U_o1qty", rowCount - 2, Convert.ToString(r[8].Value));    //출고수량
					oDS_PS_MM200L.SetValue("U_o1amt", rowCount - 2, Convert.ToString(r[9].Value));    //출고금액
					oDS_PS_MM200L.SetValue("U_o2qty", rowCount - 2, Convert.ToString(r[10].Value));   //타계정출고수량
					oDS_PS_MM200L.SetValue("U_o2amt", rowCount - 2, Convert.ToString(r[11].Value));   //타계정출고금액
					oDS_PS_MM200L.SetValue("U_jgqty", rowCount - 2, Convert.ToString(r[12].Value));   //재고수량
					oDS_PS_MM200L.SetValue("U_jgamt", rowCount - 2, Convert.ToString(r[13].Value));	  //재고금액

					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + (xlRow.Count - 1) + "건 Loding...!";

					for (loopCount = 1; loopCount <= columnCount; loopCount++)
					{
						System.Runtime.InteropServices.Marshal.ReleaseComObject(r[loopCount]); //메모리 해제
					}
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
				oForm.Update();

				PS_MM200_AddMatrixRow(0, false);
				sucessFlag = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox("[PH_PY105_Excel_Upload_Error]" + (char)13 + ex.Message);
			}
			finally
			{
				//액셀개체 닫음
				xlapp.Quit();

				System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRow);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCell);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsh);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(xlshs);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwb);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwbs);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp);
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				if (sucessFlag == true)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("엑셀 Loding 완료", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
				}

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
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_MM200_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}

							if (PS_MM200_MatrixSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}

							oForm.Items.Item("Code").Specific.Value = oForm.Items.Item("YEAR").Specific.Value.ToString().Trim() + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + oForm.Items.Item("Gubun").Specific.Value.ToString().Trim();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_MM200_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}

							if (PS_MM200_MatrixSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						PS_MM200_FormItemEnabled();
						PS_MM200_AddMatrixRow(0, true);
					}
					else if (pVal.ItemUID == "Btn02")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_MM200_Excel_Upload);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "ItemCode")
							{
								PS_MM200_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
		/// Raise_EVENT_MATRIX_LOAD 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
					PS_MM200_AddMatrixRow(oMat.VisualRowCount, false);
					PS_MM200_FormItemEnabled();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_RESIZE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM200H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM200L);
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
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
						case "1281": //찾기
							oForm.DataBrowser.BrowseBy = "Code";
							break;
						case "1282": //추가
							oForm.DataBrowser.BrowseBy = "Code"; //UDO방식일때
							PS_MM200_AddMatrixRow(0, true); //UDO방식
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
						case "1293": //행삭제
							if (oMat.RowCount != oMat.VisualRowCount)
							{
								for (int i = 1; i <= oMat.VisualRowCount; i++)
								{
									oMat.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
								}
								oMat.FlushToDataSource();
								// DBDataSource에 레코드가 한줄 더 생긴다.
								oDS_PS_MM200L.RemoveRecord(oDS_PS_MM200L.Size - 1);
								// 레코드 한 줄을 지운다.
								oMat.LoadFromDataSource();
								// DBDataSource를 매트릭스에 올리고
								if (oMat.RowCount == 0)
								{
									PS_MM200_AddMatrixRow(1, false);
								}
								else
								{
									if (!string.IsNullOrEmpty(oDS_PS_MM200L.GetValue("U_ItemCode", oMat.RowCount - 1).ToString().Trim()))
									{
										PS_MM200_AddMatrixRow(1, false);
									}
								}
							}
							break;
						case "1281": //찾기
							PS_MM200_AddMatrixRow(0, true);//UDO방식
							PS_MM200_FormItemEnabled();
							break;
						case "1282": //추가
							PS_MM200_FormItemEnabled(); //UDO방식
							PS_MM200_AddMatrixRow(0, true); //UDO방식
							break;
						case "1287": //복제
							oDS_PS_MM200H.SetValue("Code", 0, "");
							oDS_PS_MM200H.SetValue("U_YEAR", 0, "");

							for (int i = 0; i <= oMat.VisualRowCount - 1; i++)
							{
								oMat.FlushToDataSource();
								oDS_PS_MM200L.SetValue("Code", i, "");
								oMat.LoadFromDataSource();
							}
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							PS_MM200_FormItemEnabled();
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
						//행삭제전 행삭제가능여부검사
					}
					else if (pVal.BeforeAction == false)
					{
						for (i = 1; i <= oMat.VisualRowCount; i++)
						{
							oMat.Columns.Item("LineId").Cells.Item(i).Specific.Value = i;
						}
						oMat.FlushToDataSource();
						oDS_PS_MM200L.RemoveRecord(oDS_PS_MM200L.Size - 1);
						oMat.LoadFromDataSource();
						if (oMat.RowCount == 0)
						{
							PS_MM200_AddMatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_MM200L.GetValue("U_ItemCode", oMat.RowCount - 1).ToString().Trim()))
							{
								PS_MM200_AddMatrixRow(oMat.RowCount, false);
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}
	}
}
