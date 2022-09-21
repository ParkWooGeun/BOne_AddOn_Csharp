using System;
using SAPbouiCOM;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 자재 순환품 관리 Sub Form
	/// </summary>
	internal class PS_MM013 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_MM013L; //등록라인

		////부모폼
		private SAPbouiCOM.Form oBaseForm01;
		private string oBaseItemUID01;
		private string oBaseColUID01;
		private int oBaseColRow01;
		private int oBaseSelectedLineNum01;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oForm02"></param>
		/// <param name="oItemUID02"></param>
		/// <param name="oColUID02"></param>
		/// <param name="oColRow02"></param>
		/// <param name="SelectedLineNum"></param>
		public void LoadForm(ref SAPbouiCOM.Form oForm02, string oItemUID02, string oColUID02, int oColRow02, int SelectedLineNum)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM013.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM013_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM013");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

				oForm.Freeze(true);
				oBaseForm01 = oForm02;
				oBaseItemUID01 = oItemUID02;
				oBaseColUID01 = oColUID02;
				oBaseColRow01 = oColRow02;
				oBaseSelectedLineNum01 = SelectedLineNum;

				PS_MM013_CreateItems();
				PS_MM013_LoadData01();
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
		/// PS_MM013_CreateItems
		/// </summary>
		private void PS_MM013_CreateItems()
		{
			try
			{
				oDS_PS_MM013L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM013_LoadData01
		/// </summary>
		private void PS_MM013_LoadData01()
		{
			int i;
			string BPLID;
			string Year_Renamed;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				BPLID = oBaseForm01.Items.Item("BPLId").Specific.Value.ToString().Trim();;
				Year_Renamed = oBaseForm01.Items.Item("Year").Specific.Value.ToString().Trim();;

				sQry = "EXEC PS_MM013_01 '" + BPLID + "','" + Year_Renamed + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_MM013L.Clear();

				if (oRecordSet.RecordCount == 0)
				{
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				ProgressBar01.Text = "조회시작!";

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_MM013L.Size)
					{
						oDS_PS_MM013L.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_MM013L.Offset = i;
					oDS_PS_MM013L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_MM013L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim());
					oDS_PS_MM013L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("LineId").Value.ToString().Trim());
					oDS_PS_MM013L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("U_E_BANFN").Value.ToString().Trim());
					oDS_PS_MM013L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("U_E_BNFPO").Value.ToString().Trim());
					oDS_PS_MM013L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("DueDate").Value.ToString().Trim());
					oDS_PS_MM013L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("U_ItemCode").Value.ToString().Trim());
					oDS_PS_MM013L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("U_ItemName").Value.ToString().Trim());
					oDS_PS_MM013L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("U_OutSize").Value.ToString().Trim());
					oDS_PS_MM013L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("U_OutUnit").Value.ToString().Trim());
					oDS_PS_MM013L.SetValue("U_ColNum01", i, oRecordSet.Fields.Item("U_Weight").Value.ToString().Trim());
					oDS_PS_MM013L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("U_CardCode").Value.ToString().Trim());
					oDS_PS_MM013L.SetValue("U_ColReg11", i, oRecordSet.Fields.Item("U_CardName").Value.ToString().Trim());
					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
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
		/// PS_MM013_SetBaseForm
		/// </summary>
		private void PS_MM013_SetBaseForm()
		{
			int i;
			int j;
			SAPbouiCOM.Matrix oBaseMat01;
			
			try
			{
				oBaseForm01.Freeze(true);
				oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;

				oMat.FlushToDataSource();
				j = oBaseColRow01;

				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					if (oDS_PS_MM013L.GetValue("U_ColReg20", i).ToString().Trim() == "Y")
					{
						oBaseMat01.Columns.Item("PQDocNum").Cells.Item(j).Specific.Value = oDS_PS_MM013L.GetValue("U_ColReg01", i).ToString().Trim();
						oBaseMat01.Columns.Item("PQLinNum").Cells.Item(j).Specific.Value = oDS_PS_MM013L.GetValue("U_ColReg02", i).ToString().Trim();
						oBaseMat01.Columns.Item("E_BANFN").Cells.Item(j).Specific.Value = oDS_PS_MM013L.GetValue("U_ColReg03", i).ToString().Trim();
						oBaseMat01.Columns.Item("E_BNFPO").Cells.Item(j).Specific.Value = oDS_PS_MM013L.GetValue("U_ColReg04", i).ToString().Trim();
						oBaseMat01.Columns.Item("DueDate").Cells.Item(j).Specific.Value = oDS_PS_MM013L.GetValue("U_ColReg05", i).ToString().Trim();
						oBaseMat01.Columns.Item("ItemCode").Cells.Item(j).Specific.Value = oDS_PS_MM013L.GetValue("U_ColReg06", i).ToString().Trim();
						oBaseMat01.Columns.Item("ItemName").Cells.Item(j).Specific.Value = oDS_PS_MM013L.GetValue("U_ColReg07", i).ToString().Trim();
						oBaseMat01.Columns.Item("Size").Cells.Item(j).Specific.Value = oDS_PS_MM013L.GetValue("U_ColReg08", i).ToString().Trim();
						oBaseMat01.Columns.Item("Unit").Cells.Item(j).Specific.Value = oDS_PS_MM013L.GetValue("U_ColReg09", i).ToString().Trim();
						oBaseMat01.Columns.Item("Qty").Cells.Item(j).Specific.Value = oDS_PS_MM013L.GetValue("U_ColNum01", i).ToString().Trim();
						oBaseMat01.Columns.Item("CardCode").Cells.Item(j).Specific.Value = oDS_PS_MM013L.GetValue("U_ColReg10", i).ToString().Trim();
						oBaseMat01.Columns.Item("CardName").Cells.Item(j).Specific.Value = oDS_PS_MM013L.GetValue("U_ColReg11", i).ToString().Trim();
						oBaseMat01.Columns.Item("Mm01").Cells.Item(j).Specific.Value = "0";
						oBaseMat01.Columns.Item("Mm02").Cells.Item(j).Specific.Value = "0";
						oBaseMat01.Columns.Item("Mm03").Cells.Item(j).Specific.Value = "0";
						oBaseMat01.Columns.Item("Mm04").Cells.Item(j).Specific.Value = "0";
						oBaseMat01.Columns.Item("Mm05").Cells.Item(j).Specific.Value = "0";
						oBaseMat01.Columns.Item("Mm06").Cells.Item(j).Specific.Value = "0";
						oBaseMat01.Columns.Item("Mm07").Cells.Item(j).Specific.Value = "0";
						oBaseMat01.Columns.Item("Mm08").Cells.Item(j).Specific.Value = "0";
						oBaseMat01.Columns.Item("Mm09").Cells.Item(j).Specific.Value = "0";
						oBaseMat01.Columns.Item("Mm10").Cells.Item(j).Specific.Value = "0";
						oBaseMat01.Columns.Item("Mm11").Cells.Item(j).Specific.Value = "0";
						oBaseMat01.Columns.Item("Mm12").Cells.Item(j).Specific.Value = "0";
						oBaseMat01.Columns.Item("MmTot").Cells.Item(j).Specific.Value = "0";
						j += 1;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				oBaseForm01.Freeze(false);
				oForm.Close();
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
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                //    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Btn01")
					{
						PS_MM013_SetBaseForm();
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
			int i;
			string Chk = string.Empty;

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Mat01" && pVal.Row == 0 && pVal.ColUID == "Check")
					{
						oMat.FlushToDataSource();
						if (string.IsNullOrEmpty(oDS_PS_MM013L.GetValue("U_ColReg20", 0).ToString().Trim()) || oDS_PS_MM013L.GetValue("U_ColReg20", 0).ToString().Trim() == "N")
						{
							Chk = "Y";
						}
						else if (oDS_PS_MM013L.GetValue("U_ColReg20", 0).ToString().Trim() == "Y")
						{
							Chk = "N";
						}
						for (i = 0; i <= oMat.VisualRowCount - 1; i++)
						{
							oDS_PS_MM013L.SetValue("U_ColReg20", i, Chk);
						}
						oMat.LoadFromDataSource();
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
					oMat.AutoResizeColumns();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM013L);
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
                        //Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        //Raise_EVENT_FORM_DATA_ADD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        //Raise_EVENT_FORM_DATA_UPDATE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        //Raise_EVENT_FORM_DATA_DELETE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
