using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 검사게이지관리등록
	/// </summary>
	internal class PS_QM260 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_QM260H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_QM260L; //등록라인

		/// <summary>
		/// Form 호출
		/// </summary>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM260.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM260_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM260");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "Code";
				
				oForm.Freeze(true);
				PS_QM260_CreateItems();
				PS_QM260_ComboBox_Setting();
				PS_QM260_SetDocument(oFormDocEntry);

				oForm.EnableMenu("1293", true); // 행삭제
				oForm.EnableMenu("1287", true); // 복제
				oForm.EnableMenu("1284", true); // 취소
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
		/// PS_QM260_CreateItems
		/// </summary>
		private void PS_QM260_CreateItems()
		{
			try
			{
				oDS_PS_QM260H = oForm.DataSources.DBDataSources.Item("@PS_QM260H");
				oDS_PS_QM260L = oForm.DataSources.DBDataSources.Item("@PS_QM260L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.AutoResizeColumns();

				oForm.Items.Item("YM").Specific.Value = DateTime.Now.ToString("yyyyMM");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM260_ComboBox_Setting
		/// </summary>
		private void PS_QM260_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_QM260_SetDocument
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		private void PS_QM260_SetDocument(string oFromDocEntry01)
		{
			try
			{
				if (string.IsNullOrEmpty(oFromDocEntry01))
				{
					PS_QM260_FormItemEnabled();
					PS_QM260_AddMatrixRow(0, true);
				}
				else
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					PS_QM260_FormItemEnabled();
					oForm.Items.Item("Code").Specific.Value = oFromDocEntry01;
					oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					oForm.Items.Item("Code").Enabled = false;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM260_FormItemEnabled
		/// </summary>
		private void PS_QM260_FormItemEnabled()
		{
			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.EnableMenu("1281", true);  //찾기
					oForm.EnableMenu("1282", false); //추가
					oForm.Items.Item("YM").Enabled = true;
					oForm.Items.Item("Seq").Enabled = true;
					oForm.Items.Item("GaugeNo").Enabled = true;
					oForm.Items.Item("Spec").Enabled = true;
					oForm.Items.Item("TestLoc1").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.EnableMenu("1281", true);	//찾기
					oForm.EnableMenu("1282", true); //추가
					oForm.Items.Item("YM").Enabled = true;
					oForm.Items.Item("Seq").Enabled = true;
					oForm.Items.Item("GaugeNo").Enabled = true;
					oForm.Items.Item("Code").Enabled = true;
					oForm.Items.Item("Spec").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("YM").Enabled = true;
					oForm.Items.Item("Seq").Enabled = true;
					oForm.Items.Item("GaugeNo").Enabled = true;
					oForm.EnableMenu("1282", true); //추가
					oForm.Items.Item("Spec").Enabled = true;
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
		/// PS_QM260_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_QM260_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oRow = oMat.RowCount;
					oDS_PS_QM260L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_QM260L.Offset = oRow;
				oDS_PS_QM260L.SetValue("LineId", oRow, Convert.ToString(oRow + 1));
				oDS_PS_QM260L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
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
		/// PS_QM260_CopyMatrixRow
		/// </summary>
		private void PS_QM260_CopyMatrixRow()
		{
			int i;

			try
			{
				oDS_PS_QM260H.SetValue("Code", 0, "");
				oDS_PS_QM260H.SetValue("U_Seq", 0, "");

				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					oMat.FlushToDataSource();
					oDS_PS_QM260L.SetValue("Code", i, "");
					oMat.LoadFromDataSource();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);

			}
		}

		/// <summary>
		/// PS_QM260_MatrixSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_QM260_MatrixSpaceLineDel()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;
			int i;

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
					if (string.IsNullOrEmpty(oDS_PS_QM260L.GetValue("U_HisDate", 0).ToString().Trim()))
					{
						errMessage = "라인데이타가 없습니다. 확인하세요.";
						throw new Exception();
					}
				}

				if (oMat.VisualRowCount > 0)
				{
					for (i = 0; i <= oMat.VisualRowCount - 2; i++)
					{
						oDS_PS_QM260L.Offset = i;
						if (string.IsNullOrEmpty(oDS_PS_QM260L.GetValue("U_HisDate", i).ToString().Trim()))
                        {
							errMessage = "이력일자는 필수입력사항입니다. 확인하세요.";
							throw new Exception();
						}
						if (string.IsNullOrEmpty(oDS_PS_QM260L.GetValue("U_Location", i).ToString().Trim()))
                        {
							errMessage = "사용위치는 필수입력사항입니다. 확인하세요.";
							throw new Exception();
						}
					}

					if (string.IsNullOrEmpty(oDS_PS_QM260L.GetValue("U_HisDate", oMat.VisualRowCount - 1).ToString().Trim()))
					{
						oDS_PS_QM260L.RemoveRecord(oMat.VisualRowCount - 1);
					}
				}

				oMat.LoadFromDataSource();
				ReturnValue = true;
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);

				}
			}
			return ReturnValue;
		}

		/// <summary>
		/// PS_QM260_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_QM260_HeaderSpaceLineDel()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;

			try
			{ 
				if (string.IsNullOrEmpty(oDS_PS_QM260H.GetValue("U_GaugeNo", 0).ToString().Trim()))
                {
					errMessage = "게이지번호는 필수입력 사항입니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_QM260H.GetValue("U_Spec", 0).ToString().Trim()))
				{
					errMessage = "치수는 필수입력 사항입니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_QM260H.GetValue("U_TestLoc1", 0).ToString().Trim()))
                {
					errMessage = "검사부위명은 필수입력 사항입니다.";
					throw new Exception();
				}

				ReturnValue = true;
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			return ReturnValue;
		}

		/// <summary>
		/// PS_QM260_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_QM260_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string SerialNo;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (oUID == "ItemCode")
				{
					sQry = "Select ItemName From OITM Where ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
					oRecordSet.DoQuery(sQry);
					oForm.Items.Item("ItemName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
				}
				if (oUID == "YM")
				{
					SerialNo = oForm.Items.Item("YM").Specific.Value.ToString().Trim().Substring(2, 2) + "-" + oForm.Items.Item("YM").Specific.Value.ToString().Trim().Substring(4, 2) + "-" + oForm.Items.Item("Seq").Specific.Value.ToString().Trim();
					oForm.Items.Item("SerialNo").Specific.Value = SerialNo;
				}
				if (oUID == "Seq")
				{
					SerialNo = oForm.Items.Item("YM").Specific.Value.ToString().Trim().Substring(2, 2) + "-" + oForm.Items.Item("YM").Specific.Value.ToString().Trim().Substring(4, 2) + "-" + oForm.Items.Item("Seq").Specific.Value.ToString().Trim();
					oForm.Items.Item("SerialNo").Specific.Value = SerialNo;
				}

				if (oUID == "Mat01")
				{
					oMat.FlushToDataSource();
					switch (oCol)
					{
						case "HisDate":
							oMat.LoadFromDataSource();
							if (oRow == oMat.RowCount && !string.IsNullOrEmpty(oDS_PS_QM260L.GetValue("U_HisDate", oRow - 1).ToString().Trim()))
							{
								PS_QM260_AddMatrixRow(0, false);
							}
							oMat.Columns.Item("HisDate").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
                if (pVal.BeforeAction == true)
                {
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_QM260_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_QM260_MatrixSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if ((oForm.Items.Item("YM").Specific.Value.ToString().Trim() + oForm.Items.Item("Seq").Specific.Value.ToString().Trim()).Length != 8)
							{
								errMessage = "년월, 순번의 자리수가 8자리여야 합니다.";
								throw new Exception();
							}

							sQry = "Select Right('00000000' + Convert(Nvarchar(10),AutoKey),8) From ONNM Where ObjectCode = 'PS_QM260' ";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("Code").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_QM260_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}

							if (PS_QM260_MatrixSpaceLineDel() == false)
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
						PS_QM260_FormItemEnabled();
						PS_QM260_AddMatrixRow(0, true);
					}
					if (pVal.BeforeAction == false && pVal.ItemChanged == true)
					{
						if (pVal.ColUID == "HisDate")
						{
							PS_QM260_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
					}
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
            {
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
            try
            {
				if (pVal.Before_Action == true)
				{
					if (pVal.ItemUID == "ItemCode")
					{
						if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()))
						{
							PSH_Globals.SBO_Application.ActivateMenuItem("7425");
							BubbleEvent = false;
						}
					}
				}
				else if (pVal.Before_Action == false)
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
			try
			{
				oForm.Freeze(true);
				if (pVal.Before_Action == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ColUID == "HisDate")
						{
							PS_QM260_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
					}
				}
				else if (pVal.Before_Action == false)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "ItemCode")
						{
							PS_QM260_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
						else if (pVal.ItemUID == "YM")
						{
							PS_QM260_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
						else if (pVal.ItemUID == "Seq")
						{
							PS_QM260_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}

						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "HisDate") //이력일자가 바뀌면 한 줄 더 생기도록
							{
								PS_QM260_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
							}
						}
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
        /// MATRIX_LOAD 이벤트
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
					PS_QM260_AddMatrixRow(oMat.VisualRowCount, false);
					PS_QM260_FormItemEnabled();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM260H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM260L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// EVENT_ROW_DELETE
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
			int i;
            try
            {
				if (oMat.RowCount != oMat.VisualRowCount)
				{
					for (i = 1; i <= oMat.VisualRowCount; i++)
					{
						oMat.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
					}
					oMat.FlushToDataSource();
					oDS_PS_QM260L.RemoveRecord(oDS_PS_QM260L.Size - 1);
					oMat.LoadFromDataSource();
					if (oMat.RowCount == 0)
					{
						PS_QM260_AddMatrixRow(1, false);
					}
					else
					{
						if (!string.IsNullOrEmpty(oDS_PS_QM260L.GetValue("U_HisDate", oMat.RowCount - 1).ToString().Trim()))
						{
							PS_QM260_AddMatrixRow(oMat.RowCount, false);
						}
					}
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
							oForm.DataBrowser.BrowseBy = "Code";
							break;
                        case "1282": //추가
							oForm.DataBrowser.BrowseBy = "Code";
							break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
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
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
							PS_QM260_AddMatrixRow(0, true);
							PS_QM260_FormItemEnabled();
							break;
                        case "1282": //추가
							PS_QM260_FormItemEnabled();
							PS_QM260_AddMatrixRow(0, true);
							break;
						case "1287": //복제
							PS_QM260_CopyMatrixRow();
							break;
						case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
							PS_QM260_FormItemEnabled();
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
