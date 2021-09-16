using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 멀티검사사양서등록(신)
	/// </summary>
	internal class PS_QM600 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_QM600H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_QM600L; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oSeq;
		private string TmpCode;

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM600.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM600_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM600");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
                PS_QM600_CreateItems();
                PS_QM600_SetComboBox();
                PS_QM600_SetDocEntry();
                PS_QM600_AddMatrixRow(0, true);

                oForm.EnableMenu("1283", true); // 제거
				oForm.EnableMenu("1293", true); // 행삭제
				oForm.EnableMenu("1287", true); // 복제
				oForm.EnableMenu("1284", false); // 취소
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
				oForm.Freeze(false);
			}
		}

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_QM600_CreateItems()
        {
            try
            {
                oDS_PS_QM600H = oForm.DataSources.DBDataSources.Item("@PS_QM600H");
                oDS_PS_QM600L = oForm.DataSources.DBDataSources.Item("@PS_QM600L");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 콤보에 기본값설정
        /// </summary>
        private void PS_QM600_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", dataHelpClass.User_BPLID(), false, false);

                oForm.Items.Item("Ts_Gbn").Specific.ValidValues.Add("10", "Kgf/m2");
                oForm.Items.Item("Ts_Gbn").Specific.ValidValues.Add("20", "N/mm2");

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("InspItem"), "select U_Minor, U_CdName from [@PS_SY001L] where Code = 'Q600' AND U_UseYN = 'Y' ORDER BY U_Seq", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("InspItNm"), "select U_Minor, U_CdName from [@PS_SY001L] where Code = 'Q600' AND U_UseYN = 'Y' ORDER BY U_Seq", "", "");

                oMat01.Columns.Item("UseYN").ValidValues.Add("", "");
                oMat01.Columns.Item("UseYN").ValidValues.Add("Y", "Y");
                oMat01.Columns.Item("UseYN").ValidValues.Add("N", "N");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_QM600_SetDocEntry()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                string DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM600'", "");
                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.Value = "1";
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
        /// 메트릭스 행추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_QM600_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                if (RowIserted == false)
                {
                    oDS_PS_QM600L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_QM600L.Offset = oRow;
                oDS_PS_QM600L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 필수입력사항 체크(헤더)
        /// </summary>
        /// <returns></returns>
        private bool PS_QM600_DelHeaderSpaceLine()
        {
            bool returnValue = false;
            string CardCode;
            string ItemCode;
            string CardSeq;
            string sQry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_QM600H.GetValue("U_ItemCode", 0)))
                {
                    errMessage = "품목코드는 필수입력 사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_QM600H.GetValue("U_CardCode", 0)))
                {
                    errMessage = "거래처는 필수입력 사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_QM600H.GetValue("U_CardSeq", 0)))
                {
                    errMessage = "거래처순번은 필수입력입니다. 확인하세요.";
                    throw new Exception();
                }

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    ItemCode = oDS_PS_QM600H.GetValue("U_ItemCode", 0).ToString().Trim();
                    CardCode = oDS_PS_QM600H.GetValue("U_CardCode", 0).ToString().Trim();
                    CardSeq = oDS_PS_QM600H.GetValue("U_CardSeq", 0).ToString().Trim();
                    sQry = "Select Count(*) From [@PS_QM600H] Where U_ItemCode = '" + ItemCode + "' And U_CardCode = '" + CardCode + "' and U_CardSeq = '" + CardSeq + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (Convert.ToInt16(oRecordSet01.Fields.Item(0).Value) >= 1)
                    {
                        errMessage = "이미 등록된 자료입니다. 확인하세요.";
                        throw new Exception();
                    }
                }

                returnValue = true;
            }
            catch(Exception ex)
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 필수입력사항 체크(라인)
        /// </summary>
        /// <returns></returns>
        private bool PS_QM600_DelMatrixSpaceLine()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            
            try
            {
                oMat01.FlushToDataSource();

                if (oMat01.VisualRowCount == 1)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                oDS_PS_QM600L.RemoveRecord(oMat01.VisualRowCount - 1);
                oMat01.LoadFromDataSource();
                
                returnValue = true;
            }
            catch(Exception ex)
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
            }
            
            return returnValue;
        }










        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string sQry = null;
        //	short ErrNum = 0;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;

        //	object ChildForm01 = null;
        //	ChildForm01 = new PS_SM010();

        //	string ItemCode = null;

        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	string Minor = null;
        //	////BeforeAction = True
        //	if ((pVal.BeforeAction == true)) {
        //		switch (pVal.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //				////1
        //				if (pVal.ItemUID == "1") {
        //					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //						if (PS_QM600_DelHeaderSpaceLine() == false) {
        //							BubbleEvent = false;
        //							return;
        //						}

        //						if (PS_QM600_DelMatrixSpaceLine() == false) {
        //							BubbleEvent = false;
        //							return;
        //						}

        //						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //							PS_QM600_SetDocEntry();
        //							//// Input Code, Name

        //						}
        //					}
        //				}

        //				if (pVal.ItemUID == "Btn_Prt") {

        //					PS_QM600_PrintReport01();

        //				}
        //				break;

        //			case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //				////2
        //				// 제품코드
        //				if (pVal.CharPressed == 9) {
        //					if (pVal.ItemUID == "ItemCode") {
        //						//UPGRADE_WARNING: oForm.Items(ItemCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value)) {
        //							SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //							BubbleEvent = false;
        //						}
        //					}
        //					if (pVal.ItemUID == "CardCode") {
        //						//UPGRADE_WARNING: oForm.Items(CardCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value)) {
        //							SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //							BubbleEvent = false;
        //						}
        //					}
        //				}
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //				////5
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_CLICK:
        //				////6
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //				////7
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //				////8
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //				////10
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //				////11
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //				////18
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //				////19
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //				////20
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //				////27
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //				////3
        //				oLast_Item_UID = pVal.ItemUID;
        //				break;

        //			case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //				////4
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //				////17
        //				break;
        //		}
        //	////BeforeAction = False
        //	} else if ((pVal.BeforeAction == false)) {
        //		////메트릭스에 데이터 로드
        //		switch (pVal.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //				////1
        //				if (pVal.ItemUID == "1") {

        //					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pVal.Action_Success == true) {
        //						oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
        //						SubMain.Sbo_Application.ActivateMenuItem("1282");

        //					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pVal.Action_Success == false) {
        //						PS_QM600_EnableFormItem();
        //						PS_QM600_AddMatrixRow(1, oMat01.RowCount, ref true);
        //					}



        //				}
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //				////2
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //				////5

        //				oForm.Freeze(true);
        //				if (pVal.ItemChanged == true) {
        //					if ((pVal.ItemUID == "Mat01")) {
        //						if (pVal.ColUID == "InspItNm") {


        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							Minor = oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value;
        //							sQry = "Select U_RelCd From [@PS_SY001L] Where Code = 'Q600' And U_Minor = '" + Minor + "'";
        //							oRecordSet01.DoQuery(sQry);

        //							//UPGRADE_WARNING: oMat01.Columns(InspSpec).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oRecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oMat01.Columns.Item("InspSpec").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item(0).Value;
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oMat01.Columns.Item("UseYN").Cells.Item(pVal.Row).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //							oMat01.Columns.Item("InspSpec").Cells.Item(pVal.Row).Click();
        //						}
        //					}
        //				}

        //				//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oRecordSet01 = null;
        //				oForm.Freeze(false);
        //				break;



        //			case SAPbouiCOM.BoEventTypes.et_CLICK:
        //				////6
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //				////7
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //				////8
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //				////10

        //				oForm.Freeze(true);
        //				if (pVal.ItemChanged == true) {
        //					// 제품코드
        //					if (pVal.ItemUID == "ItemCode" | pVal.ItemUID == "CardCode") {
        //						PS_QM600_FlushToItemValue(pVal.ItemUID);

        //					} else if ((pVal.ItemUID == "Mat01")) {
        //						oMat01.FlushToDataSource();
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_QM600L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //						if (oMat01.RowCount == pVal.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_QM600L.GetValue("U_" + pVal.ColUID, pVal.Row - 1)))) {
        //							PS_QM600_AddMatrixRow(1, oMat01.VisualRowCount, ref true);
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_QM600L.SetValue("U_Seqno", pVal.Row - 1, oMat01.Columns.Item("LineNum").Cells.Item(pVal.Row).Specific.Value);
        //						}

        //						oMat01.LoadFromDataSource();
        //						oMat01.AutoResizeColumns();
        //						oForm.Update();
        //						oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
        //					}


        //				}

        //				//                If oMat01.RowCount = pVal.Row And Trim(oDS_PS_QM600L.GetValue("U_" & pVal.ColUID, pVal.Row - 1)) <> "" Then
        //				//                   PS_QM600_AddMatrixRow 1, oMat01.VisualRowCount, True
        //				//                End If
        //				oForm.Freeze(false);
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //				////11
        //				PS_QM600_AddMatrixRow(1, oMat01.VisualRowCount, ref true);
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //				////18
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //				////19
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //				////20
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //				////27
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //				////3
        //				oLast_Item_UID = pVal.ItemUID;
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //				////4
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //				////17
        //				SubMain.RemoveForms(oFormUniqueID);
        //				//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oForm = null;
        //				//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oMat01 = null;
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_ItemEvent_Error:
        //	///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //	oForm.Freeze(false);
        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;

        //	if (ErrNum == 1) {
        //		MDC_Com.MDC_GF_Message(ref "신규(추가)모드에서는 적용할 수 없습니다. 제품조회 후 처리하세요.", ref "E");
        //	} else {
        //		SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	}
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	////BeforeAction = True
        //	if ((pVal.BeforeAction == true)) {
        //		switch (pVal.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1281":
        //				//찾기
        //				break;
        //			case "1282":
        //				//추가
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				break;
        //			case "1293":
        //				//행삭제
        //				break;
        //			case "1283":
        //				//제거
        //				if (SubMain.Sbo_Application.MessageBox("문서를 제거(삭제) 하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1")) {
        //				} else {
        //					BubbleEvent = false;
        //					return;
        //				}
        //				break;
        //		}

        //	////BeforeAction = False
        //	} else if ((pVal.BeforeAction == false)) {
        //		switch (pVal.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1281":
        //				//찾기
        //				PS_QM600_EnableFormItem();
        //				break;
        //			case "1282":
        //				//추가
        //				PS_QM600_EnableFormItem();
        //				PS_QM600_SetDocEntry();
        //				PS_QM600_AddMatrixRow(0, oMat01.RowCount, ref true);
        //				oForm.Items.Item("ItemCode").Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
        //				break;

        //			case "1287":
        //				//복제
        //				oForm.Freeze(true);
        //				PS_QM600_SetDocEntry();
        //				oForm.Items.Item("ItemCode").Enabled = true;
        //				for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //					oMat01.FlushToDataSource();
        //					oDS_PS_QM600L.SetValue("DocEntry", i, "");
        //					oMat01.LoadFromDataSource();
        //				}

        //				oForm.Freeze(false);
        //				break;

        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				break;

        //			case "1293":
        //				//행삭제
        //				if (oMat01.RowCount != oMat01.VisualRowCount) {
        //					for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
        //					}
        //					oMat01.FlushToDataSource();
        //					// DBDataSource에 레코드가 한줄 더 생긴다.
        //					oDS_PS_QM600L.RemoveRecord(oDS_PS_QM600L.Size - 1);
        //					// 레코드 한 줄을 지운다.
        //					oMat01.LoadFromDataSource();
        //					// DBDataSource를 매트릭스에 올리고
        //					if (oMat01.RowCount == 0) {
        //						PS_QM600_AddMatrixRow(1, 0, ref true);
        //					} else {
        //						if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_QM600L.GetValue("U_InspItem", oMat01.RowCount - 1)))) {
        //							PS_QM600_AddMatrixRow(1, oMat01.RowCount, ref true);

        //						}
        //					}
        //				}
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_MenuEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////BeforeAction = True
        //	if ((BusinessObjectInfo.BeforeAction == true)) {
        //		switch (BusinessObjectInfo.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				PS_QM600_EnableFormItem();
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;
        //		}
        //	////BeforeAction = False
        //	} else if ((BusinessObjectInfo.BeforeAction == false)) {
        //		switch (BusinessObjectInfo.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_FormDataEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if ((eventInfo.BeforeAction == true)) {
        //		////작업
        //	} else if ((eventInfo.BeforeAction == false)) {
        //		////작업
        //	}
        //	return;
        //	Raise_RightClickEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion









        #region PS_QM600_FlushToItemValue
        //private void PS_QM600_FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
        //{
        //	string i = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	string sQry = null;
        //	string MItemCod = null;
        //	decimal Qty = default(decimal);
        //	decimal Calculate_Weight = default(decimal);
        //	string vReturnValue = null;

        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	switch (oUID) {
        //		case "ItemCode":
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			sQry = "Select ItemName, FrgnName, U_Spec1 From OITM Where ItemCode = '" + Strings.Trim(oForm.Items.Item("ItemCode").Specific.Value) + "'";
        //			oRecordSet01.DoQuery(sQry);
        //			oDS_PS_QM600H.SetValue("U_ItemName", 0, Strings.Trim(oRecordSet01.Fields.Item(0).Value));
        //			oDS_PS_QM600H.SetValue("U_FrgnName", 0, Strings.Trim(oRecordSet01.Fields.Item(1).Value));
        //			oDS_PS_QM600H.SetValue("U_Size", 0, Strings.Trim(oRecordSet01.Fields.Item(2).Value));
        //			break;

        //		case "CardCode":
        //			sQry = "select cardname from ocrd where cardtype='C' and cardcode = '" + Strings.Trim(oDS_PS_QM600H.GetValue("U_CardCode", 0)) + "'";
        //			oRecordSet01.DoQuery(sQry);
        //			oDS_PS_QM600H.SetValue("U_CardName", 0, Strings.Trim(oRecordSet01.Fields.Item(0).Value));
        //			break;

        //	}
        //	oForm.Freeze(false);
        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //}
        #endregion

        #region PS_QM600_PrintReport01
        //private void PS_QM600_PrintReport01()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	string WinTitle = null;
        //	string ReportName = null;
        //	string sQry = null;
        //	string Sub_sQry = null;
        //	int i = 0;
        //	string BPLId = null;
        //	string CardCode = null;
        //	string CardSeq = null;
        //	string ItemCode = null;

        //	SAPbobsCOM.Recordset oRecordSet = null;
        //	oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	MDC_PS_Common.ConnectODBC();

        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	BPLId = Strings.Trim(oForm.Items.Item("BPLId").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	CardCode = Strings.Trim(oForm.Items.Item("CardCode").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	CardSeq = Strings.Trim(oForm.Items.Item("CardSeq").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ItemCode = Strings.Trim(oForm.Items.Item("ItemCode").Specific.Value);

        //	WinTitle = "[PS_QM600_01] 검사규격 출력";

        //	ReportName = "PS_QM600_01.rpt";
        //	MDC_Globals.gRpt_Formula = new string[2];
        //	MDC_Globals.gRpt_Formula_Value = new string[2];


        //	MDC_Globals.gRpt_Formula[1] = "BPLId";
        //	sQry = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" + BPLId + "'";
        //	oRecordSet.DoQuery(sQry);
        //	//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	MDC_Globals.gRpt_Formula_Value[1] = oRecordSet.Fields.Item(0).Value;
        //	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet = null;
        //	MDC_Globals.gRpt_SRptSqry = new string[2];
        //	MDC_Globals.gRpt_SRptName = new string[2];
        //	MDC_Globals.gRpt_SFormula = new string[2, 2];
        //	MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

        //	//// SubReport


        //	MDC_Globals.gRpt_SFormula[1, 1] = "";
        //	MDC_Globals.gRpt_SFormula_Value[1, 1] = "";


        //	sQry = "EXEC PS_QM600_01 '" + BPLId + "','" + CardCode + "','" + CardSeq + "', '" + ItemCode + "'";


        //	if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") == false) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	}
        //	return;
        //	PS_QM600_PrintReport01_Error:
        //	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet = null;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_QM600_PrintReport01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion
    }
}
