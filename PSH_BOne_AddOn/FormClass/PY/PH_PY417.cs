using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 연말정산 은행파일생성
    /// </summary>
    internal class PH_PY417 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY417B; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY417.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY417_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY417");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY417_CreateItems();
                PH_PY417_SetDocument(oFormDocEntry);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private void PH_PY417_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oDS_PH_PY417B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                // 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 년도
                oForm.DataSources.UserDataSources.Add("StdYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("StdYear").Specific.DataBind.SetBound(true, "", "StdYear");
                oForm.Items.Item("StdYear").Specific.Value = DateTime.Now.ToString("yyyy");

                // 기준년월(급여)
                oForm.DataSources.UserDataSources.Add("YYYYMM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("YYYYMM").Specific.DataBind.SetBound(true, "", "YYYYMM");

                // 기준일(급여)
                oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");

                // 환급/징수
                oForm.DataSources.UserDataSources.Add("Div", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("Div").Specific.DataBind.SetBound(true, "", "Div");
                oForm.Items.Item("Div").Specific.ValidValues.Add("00", "전체");
                oForm.Items.Item("Div").Specific.ValidValues.Add("01", "환급");
                oForm.Items.Item("Div").Specific.ValidValues.Add("02", "징수");
                oForm.Items.Item("Div").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Div").DisplayDesc = true;

                // 소득세계
                oForm.DataSources.UserDataSources.Add("STot", SAPbouiCOM.BoDataType.dt_SUM, 50);
                oForm.Items.Item("STot").Specific.DataBind.SetBound(true, "", "STot");

                // 주민세계
                oForm.DataSources.UserDataSources.Add("JTot", SAPbouiCOM.BoDataType.dt_SUM, 50);
                oForm.Items.Item("JTot").Specific.DataBind.SetBound(true, "", "JTot");

                // 농특세계
                oForm.DataSources.UserDataSources.Add("NTot", SAPbouiCOM.BoDataType.dt_SUM, 50);
                oForm.Items.Item("NTot").Specific.DataBind.SetBound(true, "", "NTot");

                // 총계
                oForm.DataSources.UserDataSources.Add("Total", SAPbouiCOM.BoDataType.dt_SUM, 50);
                oForm.Items.Item("Total").Specific.DataBind.SetBound(true, "", "Total");

                // 급여변동자료적용
                oForm.DataSources.UserDataSources.Add("Check01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Check01").Specific.ValOn = "Y";
                oForm.Items.Item("Check01").Specific.ValOff = "N";
                oForm.Items.Item("Check01").Specific.DataBind.SetBound(true, "", "Check01");
                oForm.DataSources.UserDataSources.Item("Check01").Value = "N";

                // 지급년월
                oForm.DataSources.UserDataSources.Add("YM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("YM").Specific.DataBind.SetBound(true, "", "YM");

                // 지급종류
                oForm.DataSources.UserDataSources.Add("JOBTYP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("JOBTYP").Specific.DataBind.SetBound(true, "", "JOBTYP");
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("1", "급여");
                //oForm.Items.Item("JOBTYP").Specific.ValidValues.Add "2", "상여"
                oForm.Items.Item("JOBTYP").DisplayDesc = true;

                //지급구분
                oForm.DataSources.UserDataSources.Add("JOBGBN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("JOBGBN").Specific.DataBind.SetBound(true, "", "JOBGBN");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P212' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JOBGBN").Specific, "");
                oForm.Items.Item("JOBGBN").DisplayDesc = true;

                // Disable
                oForm.Items.Item("BtnPay").Enabled = false;
                oForm.Items.Item("YM").Enabled = false;
                oForm.Items.Item("JOBTYP").Enabled = false;
                oForm.Items.Item("JOBGBN").Enabled = false;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY417_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        private void PH_PY417_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PH_PY417_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY417_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY417_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY417_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    // 폼 DocEntry 세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);// 접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.Items.Item("StdYear").Specific.Value = Convert.ToString(DateTime.Now.Year - 1);//년도 세팅
                    oForm.EnableMenu("1281", true);   // 문서찾기
                    oForm.EnableMenu("1282", false);  // 문서추가

                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);// 접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", false);  // 문서찾기
                    oForm.EnableMenu("1282", true);   // 문서추가

                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false); // 접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", true); // 문서찾기
                    oForm.EnableMenu("1282", true); // 문서추가

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY417_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PH_PY417_MTX01()
        {
            int i;
            string sQry;
            short ErrNum = 0;
            double STot = 0; // 소득세계
            double JTot = 0; // 주민세계
            double NTot = 0; // 농특세계
            double Tot = 0; // 총계
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                oForm.Freeze(true);
                sQry = " EXEC [PH_PY417_01] ";
                sQry += "'" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "',";
                sQry += "'" + oForm.Items.Item("StdYear").Specific.Value.Trim() + "',";
                sQry += "'" + oForm.Items.Item("YYYYMM").Specific.Value.Trim() + "',";
                sQry += "'" + oForm.Items.Item("DocDate").Specific.Value.Trim() + "',";
                sQry += "'" + oForm.Items.Item("Div").Specific.Value.Trim() + "'";

                oRecordSet.DoQuery(sQry);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    oMat01.Clear();
                    ErrNum = 1;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PH_PY417B.InsertRecord(i);
                    }
                    oDS_PH_PY417B.Offset = i;
                    oDS_PH_PY417B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY417B.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("Div").Value);
                    oDS_PH_PY417B.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("BankCode").Value);
                    oDS_PH_PY417B.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("BankName").Value);
                    oDS_PH_PY417B.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("CntcName").Value);
                    oDS_PH_PY417B.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("AcctNo").Value);
                    oDS_PH_PY417B.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("SAmt").Value);
                    oDS_PH_PY417B.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("JAmt").Value);
                    oDS_PH_PY417B.SetValue("U_ColSum03", i, oRecordSet.Fields.Item("NAmt").Value);
                    oDS_PH_PY417B.SetValue("U_ColSum04", i, oRecordSet.Fields.Item("JSAmt").Value);
                    oDS_PH_PY417B.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("CLTCOD").Value);


                    STot += oRecordSet.Fields.Item("SAmt").Value;
                    JTot += oRecordSet.Fields.Item("JAmt").Value;
                    NTot += oRecordSet.Fields.Item("NAmt").Value;
                    Tot += oRecordSet.Fields.Item("JSAmt").Value;

                    oRecordSet.MoveNext();
                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }

                oForm.Items.Item("STot").Specific.Value = STot;
                oForm.Items.Item("JTot").Specific.Value = JTot;
                oForm.Items.Item("NTot").Specific.Value = NTot;
                oForm.Items.Item("Total").Specific.Value = Tot;

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조회 결과가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY417_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// PH_PY417_PY109_Update
        /// </summary>
        private void PH_PY417_PY109_Update()
        {
            string sQry;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string Param05;
            string Param06;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                Param02 = oForm.Items.Item("StdYear").Specific.Value.Trim();
                Param03 = oForm.Items.Item("YM").Specific.Value.Trim();
                Param04 = oForm.Items.Item("DocDate").Specific.Value.Trim();
                Param05 = oForm.Items.Item("JOBTYP").Specific.Value.Trim();
                Param06 = oForm.Items.Item("JOBGBN").Specific.Value.Trim();

                sQry = "EXEC PH_PY417_02 '" + Param01 + "','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + Param06 + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    PSH_Globals.SBO_Application.MessageBox("급여변동자료에 연말정산 징수자료를 업로드 실패했습니다.");
                }
                else
                {
                    if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value) == -1)
                    {
                        PSH_Globals.SBO_Application.MessageBox("급여변동자료가 없습니다. 확인바랍니다");
                    }
                    else if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value) == 0)
                    {
                        PSH_Globals.SBO_Application.MessageBox("급여변동자료에 연말정산 징수자료를 업로드 하지 못했습니다.확인바랍니다");
                    }
                    else if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value) > 0)
                    {
                        PSH_Globals.SBO_Application.MessageBox("연말정산 연말정산 징수자료를 업로드 했습니다. 급여변동자료를 확인하세요");
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY417_PY109_Update_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY417_PY109_Update
        /// </summary>
        private void PH_PY417_AddMatrixRow()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);
                oMat01.FlushToDataSource();
                oRow = oMat01.VisualRowCount;

                if (oMat01.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY417B.GetValue("U_LineNum", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY417B.Size <= oMat01.VisualRowCount)
                        {
                            oDS_PH_PY417B.InsertRecord(oRow);
                        }
                        oDS_PH_PY417B.Offset = oRow;
                        oDS_PH_PY417B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY417B.SetValue("U_ColReg01", oRow, "");
                        oDS_PH_PY417B.SetValue("U_ColReg02", oRow, "");
                        oDS_PH_PY417B.SetValue("U_ColReg03", oRow, "");
                        oDS_PH_PY417B.SetValue("U_ColSum01", oRow, "");
                        oDS_PH_PY417B.SetValue("U_ColSum02", oRow, "");
                        oDS_PH_PY417B.SetValue("U_ColSum03", oRow, "");
                        oMat01.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY417B.Offset = oRow - 1;
                        oDS_PH_PY417B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY417B.SetValue("U_ColReg01", oRow - 1, "");
                        oDS_PH_PY417B.SetValue("U_ColReg02", oRow - 1, "");
                        oDS_PH_PY417B.SetValue("U_ColReg03", oRow - 1, "");
                        oDS_PH_PY417B.SetValue("U_ColSum01", oRow - 1, "");
                        oDS_PH_PY417B.SetValue("U_ColSum02", oRow - 1, "");
                        oDS_PH_PY417B.SetValue("U_ColSum03", oRow - 1, "");
                        oMat01.LoadFromDataSource();
                    }
                }
                else if (oMat01.VisualRowCount == 0)
                {
                    oDS_PH_PY417B.Offset = oRow;
                    oDS_PH_PY417B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY417B.SetValue("U_ColReg01", oRow, "");
                    oDS_PH_PY417B.SetValue("U_ColReg02", oRow, "");
                    oDS_PH_PY417B.SetValue("U_ColReg03", oRow, "");
                    oDS_PH_PY417B.SetValue("U_ColSum01", oRow, "");
                    oDS_PH_PY417B.SetValue("U_ColSum02", oRow, "");
                    oDS_PH_PY417B.SetValue("U_ColSum03", oRow, "");
                    oMat01.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY417_PY109_Update_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

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

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
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
                            if (PH_PY417_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PH_PY417_MTX01();
                        }
                    }
                    else if (pVal.ItemUID == "BtnPay")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY417_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            PH_PY417_MTX01();

                            if (Convert.ToDouble(oForm.Items.Item("Total").Specific.Value) != 0)
                            {
                                PH_PY417_PY109_Update();
                            }
                            else
                            {
                                PSH_Globals.SBO_Application.SetStatusBarMessage("급여변동자료에 적용할 학자금자료가 없습니다.");
                            }
                        }
                    }
                    if (pVal.ItemUID == "Check01")
                    {
                        if (oForm.DataSources.UserDataSources.Item("Check01").Value == "Y")
                        {
                            oForm.Items.Item("YM").Enabled = true;
                            oForm.Items.Item("BtnPay").Enabled = true;
                            oForm.Items.Item("JOBTYP").Enabled = true;
                            oForm.Items.Item("JOBGBN").Enabled = true;
                            oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else
                        {
                            oForm.Items.Item("StdYear").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("YM").Enabled = false;
                            oForm.Items.Item("BtnPay").Enabled = false;
                            oForm.Items.Item("JOBTYP").Enabled = false;
                            oForm.Items.Item("JOBGBN").Enabled = false;
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        public bool PH_PY417_DataValidCheck()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.Value.Trim())) // 사업장
                {
                    errMessage = "사업장은 필수입니다.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("StdYear").Specific.Value.Trim()))// 년도
                {
                    errMessage = "년도는 필수입니다.";
                    throw new Exception();
                }
                if (oForm.DataSources.UserDataSources.Item("Check01").Value == "Y")
                {
                    if (string.IsNullOrEmpty(oForm.Items.Item("YM").Specific.Value.ToString().Trim()))
                    {
                        errMessage = "지급년월은 필수입니다.";
                        throw new Exception();
                    }
                    if (string.IsNullOrEmpty(oForm.Items.Item("JOBTYP").Specific.Value.ToString().Trim()))
                    {
                        errMessage = "지급종류는 필수입니다.";
                        throw new Exception();
                    }
                    if (string.IsNullOrEmpty(oForm.Items.Item("JOBGBN").Specific.Value.ToString().Trim()))
                    {
                        errMessage = "지급구분은 필수입니다.";
                        throw new Exception();
                    }
                }
                oMat01.FlushToDataSource(); // Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY417B.Size > 1)
                {
                    oDS_PH_PY417B.RemoveRecord(oDS_PH_PY417B.Size - 1);
                }
                oMat01.LoadFromDataSource();
                returnValue = true;
            }
            catch (Exception ex)
            {
                if(errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY417_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return returnValue;
        }

        /// <summary>
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oMat01.SelectRow(pVal.Row, true, false);
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
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
            string StdYear; //년도
            string CLTCOD; //사업장
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "StdYear":
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                                StdYear = oForm.Items.Item("StdYear").Specific.Value.Trim();

                                // 해당년도의 마지막 급여년월과 지급일자
                                sQry = "SELECT Distinct YM = U_YM, JIGBIL = Convert(char(8),U_JIGBIL,112) FROM [@PH_PY112A] WHERE U_JOBTYP = '1' And U_JOBGBN = '1' And U_CLTCOD =  '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                                sQry += "' And U_YM =  '" + oForm.Items.Item("StdYear").Specific.Value.Trim() + "12' ";
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("YYYYMM").Specific.Value = oRecordSet.Fields.Item("YM").Value.Trim();
                                oForm.Items.Item("DocDate").Specific.Value = oRecordSet.Fields.Item("JIGBIL").Value.Trim();
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    oMat01.LoadFromDataSource();
                    PH_PY417_FormItemEnabled();
                    PH_PY417_AddMatrixRow();
                    oMat01.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY417B);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            int i = 0;

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
                        case "7169":
                            //엑셀 내보내기
                            //엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                            PH_PY417_AddMatrixRow();
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY417_FormItemEnabled();
                            PH_PY417_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281":
                            //문서찾기
                            PH_PY417_FormItemEnabled();
                            PH_PY417_AddMatrixRow();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282":
                            //문서추가
                            PH_PY417_FormItemEnabled();
                            PH_PY417_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY417_FormItemEnabled();
                            break;
                        case "1293":
                            // 행삭제

                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                oMat01.FlushToDataSource();

                                while (i <= oDS_PH_PY417B.Size - 1)
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY417B.GetValue("U_LineNum", i)))
                                    {
                                        oDS_PH_PY417B.RemoveRecord(i);
                                        i = 0;
                                    }
                                    else
                                    {
                                        i += 1;
                                    }
                                }
                                for (i = 0; i <= oDS_PH_PY417B.Size; i++)
                                {
                                    oDS_PH_PY417B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                }
                                oMat01.LoadFromDataSource();
                            }
                            PH_PY417_AddMatrixRow();
                            break;

                        case "7169":
                            //엑셀 내보내기

                            //엑셀 내보내기 이후 처리
                            oForm.Freeze(true);
                            oDS_PH_PY417B.RemoveRecord(oDS_PH_PY417B.Size - 1);
                            oMat01.LoadFromDataSource();
                            oForm.Freeze(false);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormMenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            //36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
    }
}
