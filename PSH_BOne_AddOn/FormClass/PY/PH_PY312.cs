using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{ 
    /// <summary>
    /// 버스요금등록
    /// </summary>
    internal class PH_PY312 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.Matrix oMat2;
        private SAPbouiCOM.DBDataSource oDS_PH_USERDS01;
        private SAPbouiCOM.DBDataSource oDS_PH_USERDS02;
        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// Form 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY312.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY312_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY312");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                PH_PY312_CreateItems();
                PH_PY312_EnableMenus();
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
        private void PH_PY312_CreateItems()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                oDS_PH_USERDS01 = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oDS_PH_USERDS02 = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");

                oMat1 = oForm.Items.Item("Mat01").Specific;
                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                oMat2 = oForm.Items.Item("Mat02").Specific;
                oMat2.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat2.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 100);
                oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");

                oForm.DataSources.UserDataSources.Add("Amt1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("Amt1").Specific.DataBind.SetBound(true, "", "Amt1");

                oForm.DataSources.UserDataSources.Add("Amt2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("Amt2").Specific.DataBind.SetBound(true, "", "Amt2");

                sQry = "SELECT MAX(DocDate) FROM YPH_PY312A";
                oRecordSet.DoQuery(sQry);

                oForm.Items.Item("DocDate").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString("yyyyMMdd");

                PH_PY312_MTX01();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY312_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY312_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                    oForm.EnableMenu("1293", false); //행삭제
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                    oForm.EnableMenu("1293", false); //행삭제
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                    oForm.EnableMenu("1293", false); //행삭제
                }                
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 매트릭스 데이터 로드 #1
        /// </summary>
        private void PH_PY312_MTX01()
        {
            int i;
            short errNum = 0;
            string sQry;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            
            try
            {
                oForm.Freeze(true);

                sQry = "EXEC PH_PY312_01";
                oRecordSet.DoQuery(sQry);

                oMat1.Clear();
                oMat1.FlushToDataSource();
                oMat1.LoadFromDataSource();

                oMat2.Clear();
                oMat2.FlushToDataSource();
                oMat2.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 1;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PH_USERDS01.InsertRecord((i));
                    }
                    oDS_PH_USERDS01.Offset = i;
                    oDS_PH_USERDS01.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_USERDS01.SetValue("U_ColReg01", i, oRecordSet.Fields.Item(0).Value);
                    oDS_PH_USERDS01.SetValue("U_ColReg02", i, oRecordSet.Fields.Item(1).Value);
                    oDS_PH_USERDS01.SetValue("U_ColReg03", i, oRecordSet.Fields.Item(2).Value);
                    oDS_PH_USERDS01.SetValue("U_ColReg04", i, oRecordSet.Fields.Item(3).Value);
                    oDS_PH_USERDS01.SetValue("U_ColReg05", i, oRecordSet.Fields.Item(4).Value);
                    oDS_PH_USERDS01.SetValue("U_ColReg06", i, oRecordSet.Fields.Item(5).Value);
                    oDS_PH_USERDS01.SetValue("U_ColSum01", i, oRecordSet.Fields.Item(6).Value);
                    oDS_PH_USERDS01.SetValue("U_ColSum02", i, oRecordSet.Fields.Item(7).Value);

                    oRecordSet.MoveNext();

                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }
                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                ProgressBar01.Stop();

                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                oForm.Update();
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 메트릭스에 데이터 로드 #2
        /// </summary>
        /// <param name="MSTCOD"></param>
        private void PH_PY312_MTX02(string MSTCOD)
        {
            int i;
            short errNum = 0;
            string sQry;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            
            try
            {
                oForm.Freeze(true);

                sQry = "EXEC PH_PY312_02 '" + MSTCOD + "'";
                oRecordSet.DoQuery(sQry);

                oMat2.Clear();
                oMat2.FlushToDataSource();
                oMat2.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 1;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PH_USERDS02.InsertRecord(i);
                    }
                    oDS_PH_USERDS02.Offset = i;
                    oDS_PH_USERDS02.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_USERDS02.SetValue("U_ColReg01", i, oRecordSet.Fields.Item(0).Value);
                    oDS_PH_USERDS02.SetValue("U_ColReg02", i, oRecordSet.Fields.Item(1).Value);
                    oDS_PH_USERDS02.SetValue("U_ColReg03", i, oRecordSet.Fields.Item(2).Value);
                    oDS_PH_USERDS02.SetValue("U_ColSum01", i, oRecordSet.Fields.Item(3).Value);
                    oDS_PH_USERDS02.SetValue("U_ColSum02", i, oRecordSet.Fields.Item(4).Value);

                    oRecordSet.MoveNext();

                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }

                oMat2.LoadFromDataSource();
                oMat2.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                ProgressBar01.Stop();

                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                oForm.Update();
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 데이터 등록 및 수정
        /// </summary>
        /// <returns></returns>
        private void PH_PY312_UPDATE()
        {
            int i;
            string sQry;
            string MSTCOD;
            string CLTCOD;
            string DocDate;
            double Amt1;
            double Amt2;
            string IsNew; //신규여부

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat2.FlushToDataSource();

                MSTCOD = oMat2.Columns.Item("MSTCOD").Cells.Item(1).Specific.Value;

                sQry = " Select Count(*) From YPH_PY312A Where MSTCOD = '" + MSTCOD + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value > 0)
                {
                    IsNew = "N";
                }
                else
                {
                    IsNew = "Y";
                }

                for (i = 1; i <= oMat2.VisualRowCount; i++)
                {
                    if (!string.IsNullOrEmpty(oMat2.Columns.Item("MSTCOD").Cells.Item(i).Specific.Value.ToString().Trim()))
                    {
                        MSTCOD = oMat2.Columns.Item("MSTCOD").Cells.Item(i).Specific.Value;
                        DocDate = oMat2.Columns.Item("DocDate").Cells.Item(i).Specific.Value.Replace("-", "");
                        Amt1 = Convert.ToDouble(oMat2.Columns.Item("Amt1").Cells.Item(i).Specific.Value);
                        Amt2 = Convert.ToDouble(oMat2.Columns.Item("Amt2").Cells.Item(i).Specific.Value);

                        if (IsNew == "N")
                        {
                            sQry = "  UPDATE  YPH_PY312A";
                            sQry += " SET     Amt1 = " + Amt1 + ",";
                            sQry += "         Amt2 = " + Amt2;
                            sQry += " WHERE   MSTCOD = '" + MSTCOD + "'";
                            sQry +="          And DocDate = '" + DocDate + "'";
                        }
                        else
                        {
                            CLTCOD = "1";

                            sQry = "  INSERT INTO YPH_PY312A";
                            sQry += " (";
                            sQry += "     CLTCOD,";
                            sQry += "     MSTCOD,";
                            sQry += "     DocDate,";
                            sQry += "     Amt1,";
                            sQry += "     Amt2";
                            sQry += " ) ";
                            sQry += " VALUES";
                            sQry +=" ('";
                            sQry += CLTCOD + "','";
                            sQry += MSTCOD + "','";
                            sQry += DocDate + "',";
                            sQry += Amt1 + ",";
                            sQry += Amt2 + ")";
                        }
                        oRecordSet.DoQuery(sQry);
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 금액 일괄적용
        /// </summary>
        /// <returns></returns>
        private void PH_PY312_APPLY()
        {
            int i;
            string sQry;
            string MSTCOD;
            string CLTCOD;
            string DocDate;
            double Amt1;
            double Amt2;
            double InAmt1;
            double InAmt2;
            
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);


            try
            {
                oMat1.FlushToDataSource();

                DocDate = oForm.Items.Item("DocDate").Specific.Value;
                Amt1 = Convert.ToDouble(oForm.Items.Item("Amt1").Specific.Value);
                Amt2 = Convert.ToDouble(oForm.Items.Item("Amt2").Specific.Value);

                CLTCOD = "1"; //창원사업장

                sQry = "Select Count(*) From YPH_PY312A Where DocDate = '" + DocDate + "'";
                oRecordSet01.DoQuery(sQry);

                if (oRecordSet01.Fields.Item(0).Value == 0)
                {
                    for (i = 1; i <= oMat1.VisualRowCount; i++) //신규일자 Insert
                    {
                        MSTCOD = oMat1.Columns.Item("MSTCOD").Cells.Item(i).Specific.Value;

                        if (oMat1.Columns.Item("InForm02").Cells.Item(i).Specific.Value == "Y")
                        {
                            InAmt1 = Amt1;
                        }
                        else
                        {
                            InAmt1 = 0;
                        }

                        if (oMat1.Columns.Item("InForm03").Cells.Item(i).Specific.Value == "Y")
                        {
                            InAmt2 = Amt2;
                        }
                        else
                        {
                            InAmt2 = 0;
                        }

                        sQry = "        INSERT INTO YPH_PY312A ";
                        sQry += " (";
                        sQry += "     CLTCOD, ";
                        sQry += "     MSTCOD, ";
                        sQry += "     DocDate, ";
                        sQry += "     Amt1, ";
                        sQry += "     Amt2";
                        sQry += " ) ";
                        sQry += " VALUES";
                        sQry += " ('";
                        sQry += CLTCOD + "','";
                        sQry += MSTCOD + "','";
                        sQry += DocDate + "',";
                        sQry += InAmt1 + ",";
                        sQry += InAmt2 + ")";

                        oRecordSet02.DoQuery(sQry);

                        ProgressBar01.Value += 1;
                        ProgressBar01.Text = ProgressBar01.Value + "/" + oMat1.VisualRowCount + "건 처리중...!";
                    }
                }
                else
                {
                    for (i = 1; i <= oMat1.VisualRowCount; i++) //기존일자에 Update
                    {
                        MSTCOD = oMat1.Columns.Item("MSTCOD").Cells.Item(i).Specific.Value;

                        if (oMat1.Columns.Item("InForm02").Cells.Item(i).Specific.Value == "Y")
                        {
                            InAmt1 = Amt1;
                        }
                        else
                        {
                            InAmt1 = 0;
                        }

                        if (oMat1.Columns.Item("InForm03").Cells.Item(i).Specific.Value == "Y")
                        {
                            InAmt2 = Amt2;
                        }
                        else
                        {
                            InAmt2 = 0;
                        }

                        sQry = "  UPDATE  YPH_PY312A";
                        sQry += " SET     Amt1 = " + InAmt1 + ",";
                        sQry += "         Amt2 = " + InAmt2;
                        sQry += " WHERE   MSTCOD = '" + MSTCOD + "'";
                        sQry += "         AND DocDate = '" + DocDate + "'";
                        oRecordSet02.DoQuery(sQry);

                        ProgressBar01.Value += 1;
                        ProgressBar01.Text = ProgressBar01.Value + "/" + oMat1.VisualRowCount + "건 처리중...!";
                    }
                }
            }
            catch (Exception ex)
            {
                ProgressBar01.Stop();
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
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

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
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

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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
            string MSTCOD;
            short errNum = 0;

            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY312_FormItemEnabled();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY312_FormItemEnabled();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                            }
                        }
                    }

                    if (pVal.ItemUID == "Button01")
                    {
                        PH_PY312_MTX01();
                    }

                    if (pVal.ItemUID == "Button02")
                    {
                        if (oMat2.VisualRowCount > 0)
                        {
                            MSTCOD = oMat2.Columns.Item("MSTCOD").Cells.Item(1).Specific.Value;
                            PH_PY312_UPDATE();
                            PH_PY312_MTX01();
                            PH_PY312_MTX02(MSTCOD);
                        }
                    }

                    if (pVal.ItemUID == "Button03")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim()))
                        {
                            errNum = 1;
                            throw new Exception();
                        }
                        else
                        {
                            if (PSH_Globals.SBO_Application.MessageBox("현재 일자와 금액으로 일괄 적용하시겠습니까? ", 2, "Yes", "No") == 2)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            PH_PY312_APPLY();
                            PH_PY312_MTX01();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("일자를 입력하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
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
                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oLastItemUID = pVal.ItemUID;
                                oLastColUID = pVal.ColUID;
                                oLastColRow = pVal.Row;
                            }
                            break;
                        default:
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = "";
                            oLastColRow = 0;
                            break;
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oMat1.SelectRow(pVal.Row, true, false);
                                if (Convert.ToDouble(oMat1.Columns.Item("Amt1").Cells.Item(pVal.Row).Specific.Value) + Convert.ToDouble(oMat1.Columns.Item("Amt2").Cells.Item(pVal.Row).Specific.Value) > 0)
                                {
                                    PH_PY312_MTX02(oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value);
                                }
                                else
                                {
                                    oMat2.Clear();
                                    oMat2.FlushToDataSource();
                                    oMat2.AddRow();
                                    oMat2.Columns.Item("MSTCOD").Cells.Item(1).Specific.Value = oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value;
                                    oMat2.Columns.Item("FullName").Cells.Item(1).Specific.Value = oMat1.Columns.Item("FullName").Cells.Item(pVal.Row).Specific.Value;
                                    oMat2.Columns.Item("DocDate").Cells.Item(1).Specific.Value = codeHelpClass.Mid(oForm.Items.Item("DocDate").Specific.Value, 1, 4) + "-" + codeHelpClass.Mid(oForm.Items.Item("DocDate").Specific.Value, 5, 2) + "-" + codeHelpClass.Mid(oForm.Items.Item("DocDate").Specific.Value, 7, 2);
                                }
                            }
                            break;
                        case "Mat02":
                            if (pVal.Row > 0)
                            {
                                oMat2.SelectRow(pVal.Row, true, false);
                            }
                            break;
                    }

                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oLastItemUID = pVal.ItemUID;
                                oLastColUID = pVal.ColUID;
                                oLastColRow = pVal.Row;
                            }
                            break;
                        default:
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = "";
                            oLastColRow = 0;
                            break;
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
                    oMat1.LoadFromDataSource();
                    PH_PY312_FormItemEnabled();
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
                    SubMain.Remove_Forms(oFormUniqueID01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat2);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_USERDS01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_USERDS02);
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
                        case "1283":
                            if (PSH_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1293":
                            break;
                        case "1281":
                            break;
                        case "1282":
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY312_FormItemEnabled();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY312_FormItemEnabled();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY312_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY312_FormItemEnabled();
                            break;
                        case "1293": // 행삭제
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
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = pVal.ColUID;
                            oLastColRow = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID = pVal.ItemUID;
                        oLastColUID = "";
                        oLastColRow = 0;
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
