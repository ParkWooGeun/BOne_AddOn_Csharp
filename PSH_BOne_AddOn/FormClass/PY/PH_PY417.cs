using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 연말정산 은행파일생성
    /// </summary>
    internal class PH_PY417 : PSH_BaseClass
    {
        public string oFormUniqueID;
        public SAPbouiCOM.Matrix oMat01;

        //public SAPbouiCOM.Form oForm;

        private SAPbouiCOM.DBDataSource oDS_PH_PY417A; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PH_PY417B; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
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

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때

                oForm.Freeze(true);
                PH_PY417_CreateItems();
                PH_PY417_EnableMenus();
                PH_PY417_SetDocument(oFromDocEntry01);

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
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
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
                oForm.Items.Item("StdYear").Specific.VALUE = DateTime.Now.ToString("yyyy");

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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY417_EnableMenus()
        {
            try
            {
                //oForm.EnableMenu("1283", false); // 삭제
                //oForm.EnableMenu("1286", false); // 닫기
                //oForm.EnableMenu("1287", false); // 복제
                //oForm.EnableMenu("1285", false); // 복원
                //oForm.EnableMenu("1284", false); // 취소
                //oForm.EnableMenu("1293", false); // 행삭제
                //oForm.EnableMenu("1281", false);
                //oForm.EnableMenu("1282", true);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY417_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY417_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PH_PY417_FormItemEnabled();
                    //Call PH_PY417_AddMatrixRow(0, True) '//UDO방식일때
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY417_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
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
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    // 폼 DocEntry 세팅
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    //년도 세팅
                    oForm.Items.Item("StdYear").Specific.VALUE = Convert.ToString(DateTime.Now.Year - 1);
                    oForm.EnableMenu("1281", true);   // 문서찾기
                    oForm.EnableMenu("1282", false);  // 문서추가

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    oForm.EnableMenu("1281", false);  // 문서찾기
                    oForm.EnableMenu("1282", true);   // 문서추가

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);
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
            int i = 0;
            string sQry = null;
            short ErrNum = 0;
            double Total = 0;     // 총계
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string CLTCOD = string.Empty; // 사업장
            string StdYear = string.Empty;
            string YYYYMM = string.Empty;
            string DocDate = string.Empty;
            string Div = string.Empty;

            double STot = 0;     // 소득세계
            double JTot = 0;     // 주민세계
            double NTot = 0;     // 농특세계
            double Tot = 0;      // 총계

            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.Trim();
                StdYear = oForm.Items.Item("StdYear").Specific.VALUE.Trim();
                YYYYMM = oForm.Items.Item("YYYYMM").Specific.VALUE.Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.VALUE.Trim();
                Div = oForm.Items.Item("Div").Specific.VALUE.Trim();

                sQry = "            EXEC [PH_PY417_01] ";
                sQry = sQry + "'" + CLTCOD + "',"; //사업장
                sQry = sQry + "'" + StdYear + "',";
                sQry = sQry + "'" + YYYYMM + "',";
                sQry = sQry + "'" + DocDate + "',";
                sQry = sQry + "'" + Div + "'";

                oRecordSet.DoQuery(sQry);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if ((oRecordSet.RecordCount == 0))
                {
                    oMat01.Clear();
                    ErrNum = 1;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PH_PY417B.InsertRecord((i));
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

                    STot = STot + oRecordSet.Fields.Item("SAmt").Value;
                    JTot = JTot + oRecordSet.Fields.Item("JAmt").Value;
                    NTot = NTot + oRecordSet.Fields.Item("NAmt").Value;
                    Tot = Tot + oRecordSet.Fields.Item("JSAmt").Value;

                    oRecordSet.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";

                }

                oForm.Items.Item("STot").Specific.VALUE = STot;
                oForm.Items.Item("JTot").Specific.VALUE = JTot;
                oForm.Items.Item("NTot").Specific.VALUE = NTot;
                oForm.Items.Item("Total").Specific.VALUE = Tot;

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();

                ProgBar01.Stop();
                oForm.Freeze(false);

            }
            catch (Exception ex)
            {
                ProgBar01.Stop();

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조회 결과가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY417_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                oForm.Freeze(false);
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
        /// /// </summary>
        private void PH_PY417_PY109_Update()
        {
            string sQry = null;

            string Param01 = null;
            string Param02 = null;
            string Param03 = null;
            string Param04 = null;

            string Param05 = null;
            string Param06 = null;
            string Param07 = null;

            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                Param01 = oForm.Items.Item("CLTCOD").Specific.VALUE.Trim();
                Param02 = oForm.Items.Item("StdYear").Specific.VALUE.Trim();
                Param03 = oForm.Items.Item("YM").Specific.VALUE.Trim();
                Param04 = oForm.Items.Item("DocDate").Specific.VALUE.Trim();
                Param05 = oForm.Items.Item("JOBTYP").Specific.VALUE.Trim();
                Param06 = oForm.Items.Item("JOBGBN").Specific.VALUE.Trim();

                sQry = "EXEC PH_PY417_02 '" + Param01 + "','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + Param06 + "'";
                oRecordSet.DoQuery(sQry);

                if ((oRecordSet.RecordCount == 0))
                {
                    PSH_Globals.SBO_Application.MessageBox("급여변동자료에 연말정산 징수자료를 업로드 실패했습니다.");
                }
                else
                {
                    if (Conversion.Val(oRecordSet.Fields.Item(0).Value) == -1)
                    {
                        PSH_Globals.SBO_Application.MessageBox("급여변동자료가 없습니다. 확인바랍니다");
                    }
                    else if (Conversion.Val(oRecordSet.Fields.Item(0).Value) == 0)
                    {
                        PSH_Globals.SBO_Application.MessageBox("급여변동자료에 연말정산 징수자료를 업로드 하지 못했습니다.확인바랍니다");
                    }
                    else if (Conversion.Val(oRecordSet.Fields.Item(0).Value) > 0)
                    {
                        PSH_Globals.SBO_Application.MessageBox("연말정산 연말정산 징수자료를 업로드 했습니다. 급여변동자료를 확인하세요");
                    }
                }

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY417_PY109_Update_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY417_PY109_Update
        /// /// </summary>
        private void PH_PY417_AddMatrixRow()
        {
            int oRow = 0;

            try
            {
                oForm.Freeze(true);

                // [Mat01]
                oMat01.FlushToDataSource();
                oRow = oMat01.VisualRowCount;

                if (oMat01.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY417B.GetValue("U_LineNum", oRow - 1))))
                    {
                        if (oDS_PH_PY417B.Size <= oMat01.VisualRowCount)
                        {
                            oDS_PH_PY417B.InsertRecord((oRow));
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

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY417_PY109_Update_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
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

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

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

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                                                             //    Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "PH_PY417")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "BtnSearch")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY417_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            // 해야할일 작업
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

                            if (Conversion.Val(oForm.Items.Item("Total").Specific.VALUE) != 0)
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
                    if (pVal.ItemUID == "PH_PY417")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                    }
                    else if (pVal.ItemUID == "CntcCode" & pVal.CharPressed == Convert.ToDouble("9"))
                    {
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_KEY_DOWN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            bool functionReturnValue = false;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                // 사업장
                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.VALUE.Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                // 년도
                if (string.IsNullOrEmpty(oForm.Items.Item("StdYear").Specific.VALUE.Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("년도는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("StdYear").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                if (oForm.DataSources.UserDataSources.Item("Check01").Value == "Y")
                {
                    if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("YM").Specific.VALUE)))
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("지급년월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        functionReturnValue = false;
                        return functionReturnValue;
                    }
                    if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("JOBTYP").Specific.VALUE)))
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("지급종류는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        oForm.Items.Item("JOBTYP").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        functionReturnValue = false;
                        return functionReturnValue;
                    }
                    if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("JOBGBN").Specific.VALUE)))
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("지급구분은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        oForm.Items.Item("JOBGBN").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        functionReturnValue = false;
                        return functionReturnValue;
                    }
                }
                    oMat01.FlushToDataSource();
                //// Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY417B.Size > 1)
                {
                    oDS_PH_PY417B.RemoveRecord((oDS_PH_PY417B.Size - 1));
                }

                oMat01.LoadFromDataSource();

                functionReturnValue = true;
                return functionReturnValue;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY417_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                functionReturnValue = false;
                return functionReturnValue;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_DOUBLE_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LINK_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            string sQry = string.Empty;
            string StdYear = string.Empty;            //년도
            string CLTCOD = String.Empty;            //사업장
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "StdYear":
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.Trim();
                                StdYear = oForm.Items.Item("StdYear").Specific.VALUE.Trim();

                                // 해당년도의 마지막 급여년월과 지급일자
                                sQry = "SELECT Distinct YM = U_YM, JIGBIL = Convert(char(8),U_JIGBIL,112) FROM [@PH_PY112A] WHERE U_JOBTYP = '1' And U_JOBGBN = '1' And U_CLTCOD =  '" + CLTCOD + "' And U_YM =  '" + StdYear + "12' ";
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("YYYYMM").Specific.VALUE = oRecordSet.Fields.Item("YM").Value.Trim();
                                oForm.Items.Item("DocDate").Specific.VALUE = oRecordSet.Fields.Item("JIGBIL").Value.Trim();
                                break;
                        }
                    }
                }
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
                oForm.Freeze(false);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    //원본 소스(VB6.0 주석처리되어 있음)
                    //If (pval.ItemUID = "ItemCode") Then
                    //  Dim oDataTable01 As SAPbouiCOM.DataTable
                    //  Set oDataTable01 = pval.SelectedObjects
                    //  oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
                    //  Set oDataTable01 = Nothing
                    //End If
                    //If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
                    //  Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY417A", "U_CardCode,U_CardName")
                    //End If
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CHOOSE_FROM_LIST_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if ((pVal.BeforeAction == true))
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
                else if ((pVal.BeforeAction == false))
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
                        //            Case "1293":
                        //                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
                        case "1281":
                            ////문서찾기
                            PH_PY417_FormItemEnabled();
                            PH_PY417_AddMatrixRow();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282":
                            ////문서추가
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
                            //// 행삭제

                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                oMat01.FlushToDataSource();

                                while ((i <= oDS_PH_PY417B.Size - 1))
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY417B.GetValue("U_LineNum", i)))
                                    {
                                        oDS_PH_PY417B.RemoveRecord((i));
                                        i = 0;
                                    }
                                    else
                                    {
                                        i = i + 1;
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
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormMenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
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
            //string sQry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
            finally
            {
            }
        }
    }
}



//using Microsoft.VisualBasic;
//using Microsoft.VisualBasic.Compatibility;
//using System;
//using System.Collections;
//using System.Data;
//using System.Diagnostics;
//using System.Drawing;
//using System.Windows.Forms;
// // ERROR: Not supported in C#: OptionDeclaration
//namespace MDC_HR_Addon
//{
//	internal class PH_PY417
//	{
//////********************************************************************************
//////  File           : PH_PY417.cls
//////  Module         : 급여관리 > 정산관리
//////  Desc           : 연말정산 은행파일생성
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		public SAPbouiCOM.Matrix oMat1;

//		private SAPbouiCOM.DBDataSource oDS_PH_PY417A;
//		private SAPbouiCOM.DBDataSource oDS_PH_PY417B;

//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY417.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY417_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY417");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			//    oForm.DataBrowser.BrowseBy = "DocEntry"

//			oForm.Freeze(true);
//			PH_PY417_CreateItems();
//			PH_PY417_EnableMenus();
//			PH_PY417_SetDocument(oFromDocEntry01);
//			//    Call PH_PY417_FormResize

//			oForm.Update();
//			oForm.Freeze(false);

//			oForm.Visible = true;
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			return;
//			LoadForm_Error:

//			oForm.Update();
//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oForm = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private bool PH_PY417_CreateItems()
//		{
//			bool functionReturnValue = false;

//			string sQry = null;
//			int i = 0;

//			SAPbouiCOM.EditText oEdit = null;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbouiCOM.Columns oColumns = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//    Set oDS_PH_PY417A = oForm.DataSources.DBDataSources("@PH_PY417A")
//			oDS_PH_PY417B = oForm.DataSources.DBDataSources("@PS_USERDS01");

//			oMat1 = oForm.Items.Item("Mat01").Specific;

//			oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
//			oMat1.AutoResizeColumns();


//			//사업장_S
//			oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
//			//사업장_E

//			//년도_S
//			oForm.DataSources.UserDataSources.Add("StdYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("StdYear").Specific.DataBind.SetBound(true, "", "StdYear");
//			//년도_E

//			//기준년월_S
//			oForm.DataSources.UserDataSources.Add("YYYYMM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("YYYYMM").Specific.DataBind.SetBound(true, "", "YYYYMM");
//			//기준년월_E

//			//기준일자_S
//			oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");
//			//기준일자_E


//			//환급/징수구분_S
//			oForm.DataSources.UserDataSources.Add("Div", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Div").Specific.DataBind.SetBound(true, "", "Div");
//			//환급/징수구분_E


//			//소득세계_S
//			oForm.DataSources.UserDataSources.Add("STot", SAPbouiCOM.BoDataType.dt_SUM, 50);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("STot").Specific.DataBind.SetBound(true, "", "STot");
//			//소득세계_E

//			//주민세계_S
//			oForm.DataSources.UserDataSources.Add("JTot", SAPbouiCOM.BoDataType.dt_SUM, 50);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("JTot").Specific.DataBind.SetBound(true, "", "JTot");
//			//주민세계_E

//			//농특세계_S
//			oForm.DataSources.UserDataSources.Add("NTot", SAPbouiCOM.BoDataType.dt_SUM, 50);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("NTot").Specific.DataBind.SetBound(true, "", "NTot");
//			//농특세계_E

//			//총계_S
//			oForm.DataSources.UserDataSources.Add("Total", SAPbouiCOM.BoDataType.dt_SUM, 50);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Total").Specific.DataBind.SetBound(true, "", "Total");
//			//총계_E


//			//지급년월_S
//			oForm.DataSources.UserDataSources.Add("YM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("YM").Specific.DataBind.SetBound(true, "", "YM");
//			//지급년월_E

//			//지급종류_S
//			oForm.DataSources.UserDataSources.Add("JOBTYP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("JOBTYP").Specific.DataBind.SetBound(true, "", "JOBTYP");
//			//지급종류_E

//			//지급구분_S
//			oForm.DataSources.UserDataSources.Add("JOBGBN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("JOBGBN").Specific.DataBind.SetBound(true, "", "JOBGBN");
//			//지급구분_E

//			SAPbouiCOM.CheckBox oChkBox = null;

//			oForm.DataSources.UserDataSources.Add("Check01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
//			oChkBox = oForm.Items.Item("Check01").Specific;
//			oChkBox.ValOn = "Y";
//			oChkBox.ValOff = "N";
//			oChkBox.DataBind.SetBound(true, , "Check01");
//			oForm.DataSources.UserDataSources.Item("Check01").Value = "N";
//			//미체크로 값을 주고 폼을 로드

//			////----------------------------------------------------------------------------------------------
//			//// 기본사항
//			////----------------------------------------------------------------------------------------------

//			//사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;

//			//분기
//			oCombo = oForm.Items.Item("Div").Specific;
//			oCombo.ValidValues.Add("00", "전체");
//			oCombo.ValidValues.Add("01", "환급");
//			oCombo.ValidValues.Add("02", "징수");

//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			oForm.Items.Item("Div").DisplayDesc = true;


//			//// 지급종류
//			oCombo = oForm.Items.Item("JOBTYP").Specific;
//			oCombo.ValidValues.Add("1", "급여");
//			//oCombo.ValidValues.Add "2", "상여"
//			oForm.Items.Item("JOBTYP").DisplayDesc = true;

//			//// 지급구분
//			oCombo = oForm.Items.Item("JOBGBN").Specific;
//			sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P212' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oForm.Items.Item("JOBGBN").DisplayDesc = true;

//			oForm.Items.Item("BtnPay").Enabled = false;
//			oForm.Items.Item("YM").Enabled = false;
//			oForm.Items.Item("JOBTYP").Enabled = false;
//			oForm.Items.Item("JOBGBN").Enabled = false;

//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return functionReturnValue;
//			PH_PY417_CreateItems_Error:

//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY417_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY417_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			//    Call oForm.EnableMenu("1283", False)         '// 삭제
//			//    Call oForm.EnableMenu("1287", True)          '// 복제
//			//'    Call oForm.EnableMenu("1286", True)         '// 닫기
//			//    Call oForm.EnableMenu("1284", True)         '// 취소
//			//    Call oForm.EnableMenu("1293", True)         '// 행삭제

//			return;
//			PH_PY417_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY417_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY417_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY417_FormItemEnabled();
//				//        Call PH_PY417_AddMatrixRow
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY417_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY417_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY417_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY417_FormItemEnabled()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			SAPbouiCOM.ComboBox oCombo = null;
//			string CLTCOD = null;

//			oForm.Freeze(true);
//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {

//				//폼 DocEntry 세팅
//				//        Call PH_PY417_FormClear

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				//년도 세팅
//				//        Call oDS_PH_PY417A.setValue("U_StdYear", 0, Format(Date, "YYYY"))
//				//UPGRADE_WARNING: oForm.Items(StdYear).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("StdYear").Specific.VALUE = Convert.ToDouble(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYY")) - 1;

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", false);
//				////문서추가

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			}
//			oForm.Freeze(false);
//			return;
//			PH_PY417_FormItemEnabled_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY417_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			int i = 0;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			string StdYear = null;
//			//년도
//			string CLTCOD = null;
//			//사업장

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1

//					if (pval.BeforeAction == true) {
//						//                If pval.ItemUID = "1" Then
//						//                    If oForm.Mode = fm_ADD_MODE Then
//						//                        If PH_PY417_DataValidCheck = False Then
//						//                            BubbleEvent = False
//						//                        End If
//						//
//						//                        '//해야할일 작업
//						//                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
//						//                        If PH_PY417_DataValidCheck = False Then
//						//                            BubbleEvent = False
//						//                        End If
//						//                        '//해야할일 작업
//						//
//						//                    ElseIf oForm.Mode = fm_OK_MODE Then
//						//                    End If
//						//                End If
//						if (pval.ItemUID == "BtnSearch") {

//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								if (PH_PY417_DataValidCheck() == false) {
//									BubbleEvent = false;
//									return;
//								}
//								//
//								//                        '//해야할일 작업
//								PH_PY417_MTX01();

//							}
//						} else if (pval.ItemUID == "BtnPay") {
//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								if (PH_PY417_DataValidCheck() == false) {
//									BubbleEvent = false;
//									return;
//								}

//								PH_PY417_MTX01();

//								//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (Conversion.Val(oForm.Items.Item("Total").Specific.VALUE) != 0) {
//									PH_PY417_PY109_Update();
//								} else {
//									MDC_Globals.Sbo_Application.SetStatusBarMessage("급여변동자료에 적용할 연말정산 징수자료가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//								}

//							}
//						}
//						if (pval.ItemUID == "Check01") {

//							if (oForm.DataSources.UserDataSources.Item("Check01").Value == "Y") {
//								oForm.Items.Item("YM").Enabled = true;
//								oForm.Items.Item("BtnPay").Enabled = true;
//								oForm.Items.Item("JOBTYP").Enabled = true;
//								oForm.Items.Item("JOBGBN").Enabled = true;
//								oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							} else {
//								oForm.Items.Item("StdYear").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								oForm.Items.Item("YM").Enabled = false;
//								oForm.Items.Item("BtnPay").Enabled = false;
//								oForm.Items.Item("JOBTYP").Enabled = false;
//								oForm.Items.Item("JOBGBN").Enabled = false;
//							}
//						}


//					} else if (pval.BeforeAction == false) {
//						//                If pval.ItemUID = "1" Then
//						//                    If oForm.Mode = fm_ADD_MODE Then
//						//                        If pval.ActionSuccess = True Then
//						//                            Call PH_PY417_FormItemEnabled
//						//                            Call PH_PY417_AddMatrixRow
//						//                        End If
//						//                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
//						//                        If pval.ActionSuccess = True Then
//						//                            Call PH_PY417_FormItemEnabled
//						//                            Call PH_PY417_AddMatrixRow
//						//                        End If
//						//                    ElseIf oForm.Mode = fm_OK_MODE Then
//						//                        If pval.ActionSuccess = True Then
//						//                            Call PH_PY417_FormItemEnabled
//						//                        End If
//						//                    End If
//						//                End If
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2

//					if (pval.BeforeAction == true) {

//						if (pval.ItemUID == "Mat01") {



//						} else if (pval.ItemUID == "CntcCode" & pval.CharPressed == Convert.ToDouble("9")) {


//						}

//					} else if (pval.Before_Action == false) {

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					switch (pval.ItemUID) {
//						case "Mat01":
//							if (pval.Row > 0) {
//								oLastItemUID = pval.ItemUID;
//								oLastColUID = pval.ColUID;
//								oLastColRow = pval.Row;
//							}
//							break;
//						default:
//							oLastItemUID = pval.ItemUID;
//							oLastColUID = "";
//							oLastColRow = 0;
//							break;
//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//					////4
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//					////5
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemChanged == true) {
//							//                    Call PH_PY417_AddMatrixRow
//							oMat1.AutoResizeColumns();
//						}
//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					if (pval.BeforeAction == true) {
//						switch (pval.ItemUID) {
//							case "Mat01":
//								if (pval.Row > 0) {
//									oMat1.SelectRow(pval.Row, true, false);
//								}
//								break;
//						}

//						switch (pval.ItemUID) {
//							case "Mat01":
//								if (pval.Row > 0) {
//									oLastItemUID = pval.ItemUID;
//									oLastColUID = pval.ColUID;
//									oLastColRow = pval.Row;
//								}
//								break;

//							default:
//								oLastItemUID = pval.ItemUID;
//								oLastColUID = "";
//								oLastColRow = 0;
//								break;
//						}
//					} else if (pval.BeforeAction == false) {

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//					////7
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//					////8
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
//					////9
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					////10
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {

//						if (pval.ItemChanged == true) {

//						}

//					} else if (pval.BeforeAction == false) {

//						if (pval.ItemChanged == true) {

//							switch (pval.ItemUID) {
//								case "StdYear":
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									StdYear = oForm.Items.Item("StdYear").Specific.VALUE;

//									////해당년도의 마지막 급여년월과 지급일자
//									sQry = "SELECT Distinct YM = U_YM, JIGBIL = Convert(char(8),U_JIGBIL,112) FROM [@PH_PY112A] WHERE U_JOBTYP = '1' And U_JOBGBN = '1' And U_CLTCOD =  '" + CLTCOD + "' And U_YM =  '" + StdYear + "12' ";
//									oRecordSet.DoQuery(sQry);
//									//UPGRADE_WARNING: oForm.Items(YYYYMM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("YYYYMM").Specific.VALUE = Strings.Trim(oRecordSet.Fields.Item("YM").Value);
//									//UPGRADE_WARNING: oForm.Items(DocDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("DocDate").Specific.VALUE = Strings.Trim(oRecordSet.Fields.Item("JIGBIL").Value);
//									break;

//							}

//						}

//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					////11
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						oMat1.LoadFromDataSource();

//						PH_PY417_FormItemEnabled();
//						PH_PY417_AddMatrixRow();
//						oMat1.AutoResizeColumns();

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
//					////12
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
//					////16
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					////17
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm = null;
//						//UPGRADE_NOTE: oDS_PH_PY417A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY417A = null;
//						//UPGRADE_NOTE: oDS_PH_PY417B 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY417B = null;

//						//UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat1 = null;

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//					////18
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//					////19
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
//					////20
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//					////21
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {

//						oMat1.AutoResizeColumns();

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
//					////22
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
//					////23
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//					////27
//					if (pval.BeforeAction == true) {

//					} else if (pval.Before_Action == false) {
//						//                If pval.ItemUID = "Code" Then
//						//                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY417A", "Code")
//						//                End If
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
//					////37
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
//					////38
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_Drag:
//					////39
//					break;

//			}

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			return;
//			Raise_FormItemEvent_Error:
//			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			oForm.Freeze((false));
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{
//			int i = 0;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short loopCount = 0;
//			double FeeTot = 0;
//			double TuiTot = 0;
//			double Total = 0;

//			oForm.Freeze(true);

//			if ((pval.BeforeAction == true)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						if (MDC_Globals.Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2) {
//							BubbleEvent = false;
//							return;
//						}
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					case "1293":
//						break;
//					case "1281":
//						break;
//					case "1282":
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						break;

//					case "7169":
//						//엑셀 내보내기

//						//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
//						PH_PY417_AddMatrixRow();
//						break;

//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY417_FormItemEnabled();
//						PH_PY417_AddMatrixRow();
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY417_FormItemEnabled();
//						PH_PY417_AddMatrixRow();
//						oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1282":
//						////문서추가
//						PH_PY417_FormItemEnabled();
//						PH_PY417_AddMatrixRow();
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY417_FormItemEnabled();
//						break;
//					case "1293":
//						//// 행삭제

//						if (oMat1.RowCount != oMat1.VisualRowCount) {
//							oMat1.FlushToDataSource();

//							while ((i <= oDS_PH_PY417B.Size - 1)) {
//								if (string.IsNullOrEmpty(oDS_PH_PY417B.GetValue("U_LineNum", i))) {
//									oDS_PH_PY417B.RemoveRecord((i));
//									i = 0;
//								} else {
//									i = i + 1;
//								}
//							}

//							for (i = 0; i <= oDS_PH_PY417B.Size; i++) {
//								oDS_PH_PY417B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//							}

//							oMat1.LoadFromDataSource();
//						}
//						PH_PY417_AddMatrixRow();
//						break;

//					case "7169":
//						//엑셀 내보내기

//						//엑셀 내보내기 이후 처리
//						oForm.Freeze(true);
//						oDS_PH_PY417B.RemoveRecord(oDS_PH_PY417B.Size - 1);
//						oMat1.LoadFromDataSource();
//						oForm.Freeze(false);
//						break;

//				}
//			}
//			oForm.Freeze(false);
//			return;
//			Raise_FormMenuEvent_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((BusinessObjectInfo.BeforeAction == true)) {
//				switch (BusinessObjectInfo.EventType) {
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//						////33
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//						////34
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//						////35
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//						////36
//						break;
//				}
//			} else if ((BusinessObjectInfo.BeforeAction == false)) {
//				switch (BusinessObjectInfo.EventType) {
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//						////33
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//						////34
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//						////35
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//						////36
//						break;
//				}
//			}
//			return;
//			Raise_FormDataEvent_Error:


//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//		}

//		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {
//			} else if (pval.BeforeAction == false) {
//			}
//			switch (pval.ItemUID) {
//				case "Mat01":
//					if (pval.Row > 0) {
//						oLastItemUID = pval.ItemUID;
//						oLastColUID = pval.ColUID;
//						oLastColRow = pval.Row;
//					}
//					break;
//				default:
//					oLastItemUID = pval.ItemUID;
//					oLastColUID = "";
//					oLastColRow = 0;
//					break;
//			}
//			return;
//			Raise_RightClickEvent_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY417_AddMatrixRow()
//		{
//			int oRow = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			////[Mat1]
//			oMat1.FlushToDataSource();
//			oRow = oMat1.VisualRowCount;

//			if (oMat1.VisualRowCount > 0) {
//				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY417B.GetValue("U_LineNum", oRow - 1)))) {
//					if (oDS_PH_PY417B.Size <= oMat1.VisualRowCount) {
//						oDS_PH_PY417B.InsertRecord((oRow));
//					}
//					oDS_PH_PY417B.Offset = oRow;
//					oDS_PH_PY417B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//					oDS_PH_PY417B.SetValue("U_ColReg01", oRow, "");
//					oDS_PH_PY417B.SetValue("U_ColReg02", oRow, "");
//					oDS_PH_PY417B.SetValue("U_ColReg03", oRow, "");
//					oDS_PH_PY417B.SetValue("U_ColSum01", oRow, "");
//					oDS_PH_PY417B.SetValue("U_ColSum02", oRow, "");
//					oDS_PH_PY417B.SetValue("U_ColSum03", oRow, "");
//					oMat1.LoadFromDataSource();
//				} else {
//					oDS_PH_PY417B.Offset = oRow - 1;
//					oDS_PH_PY417B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//					oDS_PH_PY417B.SetValue("U_ColReg01", oRow - 1, "");
//					oDS_PH_PY417B.SetValue("U_ColReg02", oRow - 1, "");
//					oDS_PH_PY417B.SetValue("U_ColReg03", oRow - 1, "");
//					oDS_PH_PY417B.SetValue("U_ColSum01", oRow - 1, "");
//					oDS_PH_PY417B.SetValue("U_ColSum02", oRow - 1, "");
//					oDS_PH_PY417B.SetValue("U_ColSum03", oRow - 1, "");
//					oMat1.LoadFromDataSource();
//				}
//			} else if (oMat1.VisualRowCount == 0) {
//				oDS_PH_PY417B.Offset = oRow;
//				oDS_PH_PY417B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//				oDS_PH_PY417B.SetValue("U_ColReg01", oRow, "");
//				oDS_PH_PY417B.SetValue("U_ColReg02", oRow, "");
//				oDS_PH_PY417B.SetValue("U_ColReg03", oRow, "");
//				oDS_PH_PY417B.SetValue("U_ColSum01", oRow, "");
//				oDS_PH_PY417B.SetValue("U_ColSum02", oRow, "");
//				oDS_PH_PY417B.SetValue("U_ColSum03", oRow, "");
//				oMat1.LoadFromDataSource();
//			}

//			oForm.Freeze(false);
//			return;
//			PH_PY417_AddMatrixRow_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY417_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY417_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY417'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//			}
//			return;
//			PH_PY417_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY417_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY417_DataValidCheck()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = false;
//			int i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//사업장
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			//년도
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("StdYear").Specific.VALUE))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("년도는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("StdYear").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}


//			if (oForm.DataSources.UserDataSources.Item("Check01").Value == "Y") {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("YM").Specific.VALUE))) {
//					MDC_Globals.Sbo_Application.SetStatusBarMessage("지급년월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//					oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					functionReturnValue = false;
//					return functionReturnValue;
//				}
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("JOBTYP").Specific.VALUE))) {
//					MDC_Globals.Sbo_Application.SetStatusBarMessage("지급종류는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//					oForm.Items.Item("JOBTYP").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					functionReturnValue = false;
//					return functionReturnValue;
//				}
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("JOBGBN").Specific.VALUE))) {
//					MDC_Globals.Sbo_Application.SetStatusBarMessage("지급구분은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//					oForm.Items.Item("JOBGBN").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					functionReturnValue = false;
//					return functionReturnValue;
//				}
//			}


//			oMat1.FlushToDataSource();
//			//// Matrix 마지막 행 삭제(DB 저장시)
//			if (oDS_PH_PY417B.Size > 1)
//				oDS_PH_PY417B.RemoveRecord((oDS_PH_PY417B.Size - 1));

//			oMat1.LoadFromDataSource();

//			functionReturnValue = true;
//			return functionReturnValue;


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			PH_PY417_DataValidCheck_Error:


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY417_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY417_MTX01()
//		{

//			////메트릭스에 데이터 로드

//			int i = 0;
//			string sQry = null;

//			string Param01 = null;
//			string Param02 = null;
//			string Param03 = null;
//			string Param04 = null;
//			string Param05 = null;


//			double STot = 0;
//			//소득세계
//			double JTot = 0;
//			//주민세계
//			double NTot = 0;
//			//농특세계
//			double Tot = 0;
//			//총계

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param01 = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param02 = oForm.Items.Item("StdYear").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param03 = oForm.Items.Item("YYYYMM").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param04 = oForm.Items.Item("DocDate").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param05 = oForm.Items.Item("Div").Specific.VALUE;


//			SAPbouiCOM.ProgressBar ProgressBar01 = null;
//			ProgressBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

//			sQry = "EXEC PH_PY417_01 '" + Param01 + "','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "'";
//			oRecordSet.DoQuery(sQry);

//			oMat1.Clear();
//			oMat1.FlushToDataSource();
//			oMat1.LoadFromDataSource();

//			if ((oRecordSet.RecordCount == 0)) {
//				oMat1.Clear();
//				goto PH_PY417_MTX01_Exit;
//			}

//			for (i = 0; i <= oRecordSet.RecordCount - 1; i++) {
//				if (i != 0) {
//					oDS_PH_PY417B.InsertRecord((i));
//				}
//				oDS_PH_PY417B.Offset = i;
//				oDS_PH_PY417B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//				oDS_PH_PY417B.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("Div").Value);
//				oDS_PH_PY417B.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("BankCode").Value);
//				oDS_PH_PY417B.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("BankName").Value);
//				oDS_PH_PY417B.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("CntcName").Value);
//				oDS_PH_PY417B.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("AcctNo").Value);
//				oDS_PH_PY417B.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("SAmt").Value);
//				oDS_PH_PY417B.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("JAmt").Value);
//				oDS_PH_PY417B.SetValue("U_ColSum03", i, oRecordSet.Fields.Item("NAmt").Value);
//				oDS_PH_PY417B.SetValue("U_ColSum04", i, oRecordSet.Fields.Item("JSAmt").Value);

//				STot = STot + oRecordSet.Fields.Item("SAmt").Value;
//				JTot = JTot + oRecordSet.Fields.Item("JAmt").Value;
//				NTot = NTot + oRecordSet.Fields.Item("NAmt").Value;
//				Tot = Tot + oRecordSet.Fields.Item("JSAmt").Value;

//				oRecordSet.MoveNext();
//				ProgressBar01.Value = ProgressBar01.Value + 1;
//				ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";

//			}

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("STot").Specific.VALUE = STot;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("JTot").Specific.VALUE = JTot;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("NTot").Specific.VALUE = NTot;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Total").Specific.VALUE = Tot;



//			oMat1.LoadFromDataSource();
//			oMat1.AutoResizeColumns();
//			oForm.Update();

//			ProgressBar01.Stop();
//			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgressBar01 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY417_MTX01_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			if ((ProgressBar01 != null)) {
//				ProgressBar01.Stop();
//			}
//			MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "W");
//			return;
//			PH_PY417_MTX01_Error:
//			ProgressBar01.Stop();
//			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgressBar01 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY417_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}
//		private void PH_PY417_PY109_Update()
//		{


//			////급여변동자료에 Update

//			int i = 0;
//			string sQry = null;

//			string Param01 = null;
//			string Param02 = null;
//			string Param03 = null;
//			string Param04 = null;

//			string Param05 = null;
//			string Param06 = null;


//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param01 = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param02 = oForm.Items.Item("StdYear").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param03 = oForm.Items.Item("YM").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param04 = oForm.Items.Item("DocDate").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param05 = oForm.Items.Item("JOBTYP").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param06 = oForm.Items.Item("JOBGBN").Specific.VALUE;

//			SAPbouiCOM.ProgressBar ProgressBar01 = null;
//			ProgressBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

//			sQry = "EXEC PH_PY417_02 '" + Param01 + "','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + Param06 + "'";
//			oRecordSet.DoQuery(sQry);

//			if ((oRecordSet.RecordCount == 0)) {
//				//        Call MDC_Com.MDC_GF_Message("급여변동자료에 학자금자료를 업로드 실패했습니다.", "W")
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("급여변동자료에 연말정산 징수자료를 업로드 실패했습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			} else {
//				if (Conversion.Val(oRecordSet.Fields.Item(0).Value) == -1) {
//					MDC_Globals.Sbo_Application.MessageBox("급여변동자료가 없습니다. 확인바랍니다");
//				} else if (Conversion.Val(oRecordSet.Fields.Item(0).Value) == 0) {
//					MDC_Globals.Sbo_Application.MessageBox("급여변동자료에 연말정산 징수자료를 업로드 하지 못했습니다.확인바랍니다");
//				} else if (Conversion.Val(oRecordSet.Fields.Item(0).Value) > 0) {
//					MDC_Globals.Sbo_Application.MessageBox("연말정산 연말정산 징수자료를 업로드 했습니다. 급여변동자료를 확인하세요");
//				}
//			}

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY417_PY109_Update_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "W");
//			return;
//			PH_PY417_PY109_Update_Error:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY417_PY109_Update_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY417_Validate(string ValidateType)
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = true;
//			object i = null;
//			int j = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY417A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY417A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY417_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY417_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY417_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY417_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}
//	}
//}
