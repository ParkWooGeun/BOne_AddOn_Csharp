using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 품목선택
	/// </summary>
	internal class PS_SM021 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_SM021L; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oMat01Row01;

        private SAPbouiCOM.Form oBaseForm01; //부모폼
        private string oBaseItemUID01;
        private string oBaseColUID01;
        private int oBaseColRow01;
        private string oBaseOrdGbn01; //작지구분
        private string oBaseInputGbn01; //투입구분
        private string oBaseBPLId;
		
		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oForm02">Mother Form</param>
		/// <param name="oItemUID02">ItemUID</param>
		/// <param name="oColUID02">ColUID</param>
		/// <param name="oColRow02">ColRowID</param>
		/// <param name="oOrdGbn02">OrdGbn</param>
		/// <param name="oInputGbn02">InputGbn</param>
		/// <param name="oBPLId">BPLID</param>
		public void LoadForm(SAPbouiCOM.Form oForm02, string oItemUID02, string oColUID02, int oColRow02, string oOrdGbn02, string oInputGbn02, string oBPLId)
		{	
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SM021.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SM021_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SM021");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				//oForm.DataBrowser.BrowseBy="DocEntry"

				oForm.Freeze(true);
				oBaseForm01 = oForm02;
				oBaseItemUID01 = oItemUID02;
				oBaseColUID01 = oColUID02;
				oBaseColRow01 = oColRow02;
				oBaseOrdGbn01 = oOrdGbn02;
				oBaseInputGbn01 = oInputGbn02;
				oBaseBPLId = oBPLId;

				PS_SM021_CreateItems();
                PS_SM021_ComboBox_Setting();
                PS_SM021_CF_ChooseFromList();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
			}
		}

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_SM021_CreateItems()
        {
            try
            {
                oDS_PS_SM021L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMat01.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

                oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

                oForm.DataSources.UserDataSources.Add("ItemGpCd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItemGpCd").Specific.DataBind.SetBound(true, "", "ItemGpCd");

                oForm.DataSources.UserDataSources.Add("ItmBsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItmBsort").Specific.DataBind.SetBound(true, "", "ItmBsort");

                oForm.DataSources.UserDataSources.Add("ItmMsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItmMsort").Specific.DataBind.SetBound(true, "", "ItmMsort");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>        
        private void PS_SM021_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItemGpCd").Specific, "SELECT ItmsGrpCod, ItmsGrpNam FROM [OITB] order by ItmsGrpCod", "", false, true);
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] order by Code", "", false, true);
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] order by U_Code", "", false, true);

                oForm.Items.Item("ItemGpCd").Specific.Select("104");

                if (oBaseOrdGbn01 == "101")
                {
                    oForm.Items.Item("ItmBsort").Specific.Select("301");
                }
                else if (oBaseOrdGbn01 == "102")
                {
                    oForm.Items.Item("ItmBsort").Specific.Select("305");
                }
                else if (oBaseOrdGbn01 == "103")
                {
                    oForm.Items.Item("ItmBsort").Specific.Select("301");
                }
                else if (oBaseOrdGbn01 == "104")
                {
                    oForm.Items.Item("ItmBsort").Specific.Select("302");
                }
                else if (oBaseOrdGbn01 == "105")
                {
                    oForm.Items.Item("ItmBsort").Specific.Select("303");
                }
                else if (oBaseOrdGbn01 == "106")
                {
                    oForm.Items.Item("ItmBsort").Specific.Select("305");
                }
                else if (oBaseOrdGbn01 == "107")
                {
                    oForm.Items.Item("ItmBsort").Specific.Select("309");
                }
                else if (oBaseOrdGbn01 == "108")
                {
                    oForm.Items.Item("ItmBsort").Specific.Select("310");
                }
                else if (oBaseOrdGbn01 == "109")
                {
                    oForm.Items.Item("ItmBsort").Specific.Select("314");
                }
                else if (oBaseOrdGbn01 == "110")
                {
                    oForm.Items.Item("ItmBsort").Specific.Select("315");
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ChooseFromList 설정
        /// </summary>
        private void PS_SM021_CF_ChooseFromList()
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            SAPbouiCOM.EditText oEdit = null;
            
            try
            {
                oEdit = oForm.Items.Item("ItemCode").Specific;
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams.ObjectType = "4";
                oCFLCreationParams.UniqueID = "CFLITEMCODE";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);

                oEdit.ChooseFromListUID = "CFLITEMCODE";
                oEdit.ChooseFromListAlias = "ItemCode";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_SM021_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SM021'", "");

                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 필수입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_SM021_DataValidCheck()
        {
            bool returnValue = false;

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_SM021_FormClear();
                }

                returnValue = true;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }

            return returnValue;
        }

        /// <summary>
        /// 메트릭스 데이터 로드
        /// </summary>
        private void PS_SM021_MTX01()
        {   
            int i;
            string Query01;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
                Param02 = oForm.Items.Item("ItemGpCd").Specific.Selected.Value.ToString().Trim();
                Param03 = oForm.Items.Item("ItmBsort").Specific.Selected.Value.ToString().Trim();
                Param04 = oForm.Items.Item("ItmMsort").Specific.Selected.Value.ToString().Trim();

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                Query01 = "EXEC PS_PP030_03 '" + oBaseOrdGbn01 + "','" + oBaseInputGbn01 + "','" + Param01 + "','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + oBaseBPLId + "'";
                RecordSet01.DoQuery(Query01);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_SM021L.InsertRecord(i);
                    }
                    oDS_PS_SM021L.Offset = i;
                    oDS_PS_SM021L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_SM021L.SetValue("U_ColReg01", i, RecordSet01.Fields.Item(0).Value);
                    oDS_PS_SM021L.SetValue("U_ColReg02", i, RecordSet01.Fields.Item(1).Value);
                    oDS_PS_SM021L.SetValue("U_ColReg03", i, RecordSet01.Fields.Item(2).Value);
                    oDS_PS_SM021L.SetValue("U_ColReg04", i, RecordSet01.Fields.Item(3).Value);
                    oDS_PS_SM021L.SetValue("U_ColQty01", i, RecordSet01.Fields.Item(4).Value);
                    oDS_PS_SM021L.SetValue("U_ColReg05", i, RecordSet01.Fields.Item(5).Value);
                    RecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();
            }
            catch(Exception ex)
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
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
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
            int i;
            string sQry;
            string CntcCode;
            string CntcName;
            SAPbouiCOM.Matrix oBaseMat01 = null;
            SAPbouiCOM.DBDataSource oBaseDS_PS_PP030L = null;
            SAPbobsCOM.Recordset oRecordSet01 = null;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Button01")
                    {
                        oBaseForm01.Freeze(true);
                        ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            sQry = "Select U_MSTCOD, lastName + firstName From [OHEM] a Inner Join [OUSR] b On a.userId = b.USERID Where USER_CODE = '" + PSH_Globals.oCompany.UserName + "'";
                            oRecordSet01.DoQuery(sQry);
                            CntcCode = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                            CntcName = oRecordSet01.Fields.Item(1).Value.ToString().Trim();
                            
                            if (codeHelpClass.Left(oBaseForm01.UniqueID, 8) == "PS_PP030") //작업지시등록에서 호출하였을 경우
                            {
                                oBaseDS_PS_PP030L = oBaseForm01.DataSources.DBDataSources.Item("@PS_PP030L");
                                oBaseMat01 = oBaseForm01.Items.Item("Mat02").Specific;

                                for (i = 1; i <= oMat01.VisualRowCount; i++)
                                {
                                    if (oMat01.IsRowSelected(i) == true)
                                    {
                                        oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.Value = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
                                        oBaseDS_PS_PP030L.SetValue("U_InputGbn", oBaseColRow01 - 1, oBaseInputGbn01);
                                        oBaseDS_PS_PP030L.SetValue("U_ItemName", oBaseColRow01 - 1, oMat01.Columns.Item("ItemName").Cells.Item(i).Specific.Value);
                                        oBaseDS_PS_PP030L.SetValue("U_ItemGpCd", oBaseColRow01 - 1, oMat01.Columns.Item("ItemGpCd").Cells.Item(i).Specific.Value);
                                        oBaseDS_PS_PP030L.SetValue("U_BatchNum", oBaseColRow01 - 1, oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value);
                                        oBaseDS_PS_PP030L.SetValue("U_Weight", oBaseColRow01 - 1, oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value);

                                        sQry = "Select BuyUnitMsr From OITM Where ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim() + "'";

                                        oRecordSet01.DoQuery(sQry);
                                        oBaseDS_PS_PP030L.SetValue("U_Unit", oBaseColRow01 - 1, oRecordSet01.Fields.Item(0).Value.ToString().Trim());

                                        if (oBaseForm01.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "105")
                                        {
                                            oBaseDS_PS_PP030L.SetValue("U_CntcCode", oBaseColRow01 - 1, CntcCode);
                                            oBaseDS_PS_PP030L.SetValue("U_CntcName", oBaseColRow01 - 1, CntcName);
                                            oBaseDS_PS_PP030L.SetValue("U_ProcType", oBaseColRow01 - 1, "10");
                                            oBaseDS_PS_PP030L.SetValue("U_DueDate", oBaseColRow01 - 1, DateTime.Now.ToString("yyyyMMdd"));
                                            oBaseDS_PS_PP030L.SetValue("U_CGDate", oBaseColRow01 - 1, DateTime.Now.ToString("yyyyMMdd"));
                                        }
                                        oBaseColRow01 += 1;
                                    }
                                }
                            }
                            else if (codeHelpClass.Left(oBaseForm01.UniqueID, 8) == "PS_PP038") //투입자재추가등록에서 호출하였을 경우
                            {
                                oBaseDS_PS_PP030L = oBaseForm01.DataSources.DBDataSources.Item("@PS_USERDS01");
                                oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;

                                for (i = 1; i <= oMat01.VisualRowCount; i++)
                                {
                                    if (oMat01.IsRowSelected(i) == true)
                                    {
                                        oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.Value = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
                                        oBaseDS_PS_PP030L.SetValue("U_ColReg02", oBaseColRow01 - 1, oBaseInputGbn01);
                                        oBaseDS_PS_PP030L.SetValue("U_ColReg04", oBaseColRow01 - 1, oMat01.Columns.Item("ItemName").Cells.Item(i).Specific.Value);
                                        oBaseDS_PS_PP030L.SetValue("U_ColReg05", oBaseColRow01 - 1, oMat01.Columns.Item("ItemGpCd").Cells.Item(i).Specific.Value);
                                        oBaseDS_PS_PP030L.SetValue("U_ColReg06", oBaseColRow01 - 1, oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value);
                                        oBaseDS_PS_PP030L.SetValue("U_ColQty01", oBaseColRow01 - 1, oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value);

                                        sQry = "Select BuyUnitMsr From OITM Where ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim() + "'";

                                        oRecordSet01.DoQuery(sQry);
                                        oBaseDS_PS_PP030L.SetValue("U_ColReg08", oBaseColRow01 - 1, oRecordSet01.Fields.Item(0).Value.ToString().Trim());

                                        if (oBaseForm01.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "105")
                                        {
                                            oBaseDS_PS_PP030L.SetValue("U_ColReg10", oBaseColRow01 - 1, CntcCode);
                                            oBaseDS_PS_PP030L.SetValue("U_ColReg11", oBaseColRow01 - 1, CntcName);
                                            oBaseDS_PS_PP030L.SetValue("U_ColReg13", oBaseColRow01 - 1, "10");
                                            oBaseDS_PS_PP030L.SetValue("U_ColDt01", oBaseColRow01 - 1, DateTime.Now.ToString("yyyyMMdd"));
                                            oBaseDS_PS_PP030L.SetValue("U_ColDt02", oBaseColRow01 - 1, DateTime.Now.ToString("yyyyMMdd"));
                                        }
                                        oBaseColRow01 += 1;
                                    }
                                }
                            }

                            oBaseMat01.LoadFromDataSource();
                            oBaseMat01.Columns.Item("Comments").Cells.Item(oBaseColRow01 - 1).Specific.Value = "비고";
                            oBaseMat01.Columns.Item("Comments").Cells.Item(oBaseColRow01 - 1).Specific.Value = "";
                            oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oBaseMat01.AutoResizeColumns();
                            oForm.Close();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Button02")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_SM021_MTX01();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
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
            finally
            {
                oBaseForm01.Freeze(false);

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

                if (oBaseMat01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBaseMat01);
                }
                
                if (oBaseDS_PS_PP030L != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBaseDS_PS_PP030L);
                }
                
                if (oRecordSet01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                }
            }
        }

        /// <summary>
        /// GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                    }
                    else
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "ItmBsort")
                    {
                        for (i = 0; i <= oForm.Items.Item("ItmMsort").Specific.ValidValues.Count - 1; i++)
                        {
                            oForm.Items.Item("ItmMsort").Specific.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        dataHelpClass.Set_ComboList(oForm.Items.Item("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] WHERE U_rCode = '" + oForm.Items.Item("ItmBsort").Specific.Selected.Value + "' ORDER BY U_Code", "", true, true);
                        if (oForm.Items.Item("ItmMsort").Specific.ValidValues.Count > 0)
                        {
                            oForm.Items.Item("ItmMsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            if (oMat01.IsRowSelected(pVal.Row) == true)
                            {
                                oMat01.SelectRow(pVal.Row, false, true);
                            }
                            else
                            {
                                oMat01.SelectRow(pVal.Row, true, true);
                            }
                        }
                        oMat01Row01 = pVal.Row;
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            string sQry;
            string CntcCode;
            string CntcName;
            SAPbouiCOM.Matrix oBaseMat01 = null;
            SAPbouiCOM.DBDataSource oBaseDS_PS_PP030L = null;
            SAPbobsCOM.Recordset oRecordSet01 = null;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.Row == 0) //헤더를 더블클릭했을 경우
                            {
                                oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true; //정렬
                                oMat01.FlushToDataSource();
                            }
                            
                            if (oMat01Row01 > 0) //매트릭스 내부행을 클릭했을 경우
                            {
                                oBaseForm01.Freeze(true);
                                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                                oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                sQry = "Select U_MSTCOD, lastName + firstName From [OHEM] a Inner Join [OUSR] b On a.userId = b.USERID Where USER_CODE = '" + PSH_Globals.oCompany.UserName + "'";
                                oRecordSet01.DoQuery(sQry);
                                CntcCode = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                                CntcName = oRecordSet01.Fields.Item(1).Value.ToString().Trim();
                                
                                if (codeHelpClass.Left(oBaseForm01.UniqueID, 8) == "PS_PP038") //투입자재추가등록에서 호출하였을 경우
                                {
                                    oBaseDS_PS_PP030L = oBaseForm01.DataSources.DBDataSources.Item("@PS_USERDS01");
                                    oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;

                                    for (i = 1; i <= oMat01.VisualRowCount; i++)
                                    {
                                        if (oMat01.IsRowSelected(i) == true)
                                        {
                                            oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.Value = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
                                            oBaseDS_PS_PP030L.SetValue("U_ColReg02", oBaseColRow01 - 1, oBaseInputGbn01);
                                            oBaseDS_PS_PP030L.SetValue("U_ColReg04", oBaseColRow01 - 1, oMat01.Columns.Item("ItemName").Cells.Item(i).Specific.Value);
                                            oBaseDS_PS_PP030L.SetValue("U_ColReg05", oBaseColRow01 - 1, oMat01.Columns.Item("ItemGpCd").Cells.Item(i).Specific.Value);
                                            oBaseDS_PS_PP030L.SetValue("U_ColReg06", oBaseColRow01 - 1, oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value);
                                            oBaseDS_PS_PP030L.SetValue("U_ColQty01", oBaseColRow01 - 1, oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value);

                                            sQry = "Select BuyUnitMsr From OITM Where ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                                            oRecordSet01.DoQuery(sQry);
                                            oBaseDS_PS_PP030L.SetValue("U_ColReg08", oBaseColRow01 - 1, oRecordSet01.Fields.Item(0).Value.ToString().Trim());

                                            if (oBaseForm01.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "105")
                                            {
                                                oBaseDS_PS_PP030L.SetValue("U_ColReg10", oBaseColRow01 - 1, CntcCode);
                                                oBaseDS_PS_PP030L.SetValue("U_ColReg11", oBaseColRow01 - 1, CntcName);
                                                oBaseDS_PS_PP030L.SetValue("U_ColReg13", oBaseColRow01 - 1, "10");
                                                oBaseDS_PS_PP030L.SetValue("U_ColDt01", oBaseColRow01 - 1, DateTime.Now.ToString("yyyyMMdd"));
                                                oBaseDS_PS_PP030L.SetValue("U_ColDt02", oBaseColRow01 - 1, DateTime.Now.ToString("yyyyMMdd"));
                                            }
                                            oBaseColRow01 += 1;
                                        }
                                    }
                                    oBaseMat01.LoadFromDataSource();
                                    oBaseMat01.Columns.Item("Comments").Cells.Item(oBaseColRow01 - 1).Specific.Value = "비고";
                                    oBaseMat01.Columns.Item("Comments").Cells.Item(oBaseColRow01 - 1).Specific.Value = "";
                                    oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    oBaseMat01.AutoResizeColumns();
                                    oForm.Close();
                                }
                            }

                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oBaseForm01.Freeze(false);

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

                if (oBaseMat01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBaseMat01);
                }

                if (oBaseDS_PS_PP030L != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBaseDS_PS_PP030L);
                }

                if (oRecordSet01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                }
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
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SM021L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            SAPbouiCOM.DataTable oDataTable01 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; //ItemEvent를 ChooseFromListEvent로 명시적 형변환 후 SelectedObjects 할당

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (oDataTable01 != null) //SelectedObjects 가 null이 아닐때만 실행(ChooseFromList 팝업창을 취소했을 때 미실행)
                    {
                        if (pVal.ItemUID == "ItemCode")
                        {
                            oForm.DataSources.UserDataSources.Item("ItemCode").Value = oDataTable01.Columns.Item("ItemCode").Cells.Item(0).Value;
                            oForm.DataSources.UserDataSources.Item("ItemName").Value = oDataTable01.Columns.Item("ItemName").Cells.Item(0).Value;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oDataTable01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDataTable01);
                }
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
            finally
            {
            }
        }

        /// <summary>
        /// RightClickEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                }
                switch (pVal.ItemUID)
                {
                    case "Mat01":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
