using System;

using SAPbouiCOM;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 배부사이클실행결과조회
    /// </summary>
    internal class PS_CO111 : PSH_BaseClass
    {
        private string oFormUniqueID;

        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        
        private SAPbouiCOM.DBDataSource oDS_PS_CO110M; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_CO110N; //등록라인

        //부모폼 정보
        private SAPbouiCOM.Form oBaseForm01;
        private string oBaseItemUID01;
        private string oBaseColUID01;
        private int oBaseColRow01;
        private int oBaseSelectedLineNum01;

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="pForm">부모폼</param>
        /// <param name="pItemUID">부모폼 ItemUID</param>
        /// <param name="pColUID">부모폼 Matrix ColUID</param>
        /// <param name="pColRow">부모폼 Matrix ColRow</param>
        /// <param name="pSelectedLineNum">부모폼 Matrix 선택된 행번호</param>
        public void LoadForm(ref SAPbouiCOM.Form pForm, string pItemUID, string pColUID, int pColRow, int pSelectedLineNum)
        {

            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO111.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_CO111_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_CO111");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);

                oBaseForm01 = pForm;
                oBaseItemUID01 = pItemUID;
                oBaseColUID01 = pColUID;
                oBaseColRow01 = pColRow;
                oBaseSelectedLineNum01 = pSelectedLineNum;

                CreateItems();
                LoadData01();
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
        /// 화면 Item 생성
        /// </summary>
        private void CreateItems()
        {
            try
            {
                oDS_PS_CO110M = oForm.DataSources.DBDataSources.Item("@PS_CO110M");
                oDS_PS_CO110N = oForm.DataSources.DBDataSources.Item("@PS_CO110N");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat02 = oForm.Items.Item("Mat02").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat02.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("RadioBtn01").Specific.DataBind.SetBound(true, "", "Opt01");
                oForm.Items.Item("RadioBtn02").Specific.DataBind.SetBound(true, "", "Opt02");
                oForm.Items.Item("RadioBtn01").Specific.GroupWith("RadioBtn02");

                oForm.DataSources.UserDataSources.Add("SumCost1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SumCost1").Specific.DataBind.SetBound(true, "", "SumCost1");

                oForm.DataSources.UserDataSources.Add("SumCost2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SumCost2").Specific.DataBind.SetBound(true, "", "SumCost2");

                oForm.DataSources.UserDataSources.Add("SumCost3", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SumCost3").Specific.DataBind.SetBound(true, "", "SumCost3");

                oForm.DataSources.UserDataSources.Add("SumCost4", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SumCost4").Specific.DataBind.SetBound(true, "", "SumCost4");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void LoadData01()
        {
            short i = 0;
            string sQry;
            string errCode = string.Empty;
            
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            try
            {
                string Code = oBaseForm01.Items.Item("Code").Specific.Value.ToString().Trim();

                sQry = "Select * From [@PS_CO110M] Where Code = '" + Code + "' And U_L_LineId = '" + oBaseSelectedLineNum01 + "'";
                oRecordSet01.DoQuery(sQry);

                oForm.Freeze(true);

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

                oMat01.Clear();
                oDS_PS_CO110M.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errCode = "1";
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_CO110M.Size)
                    {
                        oDS_PS_CO110M.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_CO110M.Offset = i;
                    oDS_PS_CO110M.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_CO110M.SetValue("U_CCCode", i, oRecordSet01.Fields.Item("U_CCCode").Value.ToString().Trim());
                    oDS_PS_CO110M.SetValue("U_CCName", i, oRecordSet01.Fields.Item("U_CCName").Value.ToString().Trim());
                    oDS_PS_CO110M.SetValue("U_CECode", i, oRecordSet01.Fields.Item("U_CECode").Value.ToString().Trim());
                    oDS_PS_CO110M.SetValue("U_CEName", i, oRecordSet01.Fields.Item("U_CEName").Value.ToString().Trim());
                    oDS_PS_CO110M.SetValue("U_Cost1", i, oRecordSet01.Fields.Item("U_Cost1").Value.ToString().Trim());
                    oDS_PS_CO110M.SetValue("U_Cost2", i, oRecordSet01.Fields.Item("U_Cost2").Value.ToString().Trim());
                    oDS_PS_CO110M.SetValue("U_Cost3", i, oRecordSet01.Fields.Item("U_Cost3").Value.ToString().Trim());
                    oDS_PS_CO110M.SetValue("U_Cost4", i, oRecordSet01.Fields.Item("U_Cost4").Value.ToString().Trim());
                    oDS_PS_CO110M.SetValue("LineId", i, oRecordSet01.Fields.Item("LineId").Value);

                    oRecordSet01.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("조회 결과가 없습니다. 확인하세요.");
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                ProgBar01.Stop();
                oForm.Freeze(false);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
            }
        }

        /// <summary>
        /// 상세 데이터 조회
        /// </summary>
        /// <param name="pLineNum"></param>
        private void LoadData02(int pLineNum)
        {
            string sQry;
            string errCode = string.Empty;
            
            double SumCost3 = 0;
            double SumCost1 = 0;
            double SumCost2 = 0;
            double SumCost4 = 0;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            try
            {
                string Code = oBaseForm01.Items.Item("Code").Specific.Value;

                sQry = "Select * From [@PS_CO110N] Where Code = '" + Code + "' And U_M_LineId = '" + pLineNum + "'";
                oRecordSet01.DoQuery(sQry);

                oForm.Freeze(true);

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

                oMat02.Clear();
                oDS_PS_CO110N.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errCode = "1";
                    throw new Exception();
                }

                for (int i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_CO110N.Size)
                    {
                        oDS_PS_CO110N.InsertRecord(i);
                    }

                    oMat02.AddRow();
                    oDS_PS_CO110N.Offset = i;
                    oDS_PS_CO110N.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_CO110N.SetValue("U_CCCode", i, oRecordSet01.Fields.Item("U_CCCode").Value.ToString().Trim());
                    oDS_PS_CO110N.SetValue("U_CCName", i, oRecordSet01.Fields.Item("U_CCName").Value.ToString().Trim());
                    oDS_PS_CO110N.SetValue("U_CECode", i, oRecordSet01.Fields.Item("U_CECode").Value.ToString().Trim());
                    oDS_PS_CO110N.SetValue("U_CEName", i, oRecordSet01.Fields.Item("U_CEName").Value.ToString().Trim());
                    oDS_PS_CO110N.SetValue("U_Cost1", i, oRecordSet01.Fields.Item("U_Cost1").Value.ToString().Trim());
                    oDS_PS_CO110N.SetValue("U_Cost2", i, oRecordSet01.Fields.Item("U_Cost2").Value.ToString().Trim());
                    oDS_PS_CO110N.SetValue("U_Cost3", i, oRecordSet01.Fields.Item("U_Cost3").Value.ToString().Trim());
                    oDS_PS_CO110N.SetValue("U_Cost4", i, oRecordSet01.Fields.Item("U_Cost4").Value.ToString().Trim());

                    SumCost1 = SumCost1 + Convert.ToDouble(oRecordSet01.Fields.Item("U_Cost1").Value.ToString().Trim());
                    SumCost2 = SumCost2 + Convert.ToDouble(oRecordSet01.Fields.Item("U_Cost2").Value.ToString().Trim());
                    SumCost3 = SumCost3 + Convert.ToDouble(oRecordSet01.Fields.Item("U_Cost3").Value.ToString().Trim());
                    SumCost4 = SumCost4 + Convert.ToDouble(oRecordSet01.Fields.Item("U_Cost4").Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

                oForm.Items.Item("SumCost1").Specific.Value = SumCost1;
                oForm.Items.Item("SumCost2").Specific.Value = SumCost2;
                oForm.Items.Item("SumCost3").Specific.Value = SumCost3;
                oForm.Items.Item("SumCost4").Specific.Value = SumCost4;

                oMat02.LoadFromDataSource();
                oMat02.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("조회 결과가 없습니다. 확인하세요.");
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                ProgBar01.Stop();
                oForm.Freeze(false);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
            }
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                switch (oUID)
                {
                    case "CardCode":
                        sQry = "Select CardName From [OCRD] Where CardCode = '" + oForm.DataSources.UserDataSources.Item("CardCode").Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);

                        oForm.DataSources.UserDataSources.Item("CardName").Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                    if (pVal.ItemUID == "Btn01")
                    {
                        oForm.Close();
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
                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }

                    if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat02.SelectRow(pVal.Row, true, false);
                        }
                    }

                    if (pVal.ItemUID == "RadioBtn01")
                    {
                        oForm.Freeze(true);
                        oForm.Settings.MatrixUID = "Mat01";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oForm.Freeze(false);
                    }

                    if (pVal.ItemUID == "RadioBtn02")
                    {
                        oForm.Freeze(true);
                        oForm.Settings.MatrixUID = "Mat02";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oForm.Freeze(false);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01" && pVal.Row != 0)
                    {
                        LoadData02(Convert.ToInt32(oMat01.Columns.Item("LineId").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()));
                    }
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
                    oMat01.AutoResizeColumns();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO110M);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO110N);
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
