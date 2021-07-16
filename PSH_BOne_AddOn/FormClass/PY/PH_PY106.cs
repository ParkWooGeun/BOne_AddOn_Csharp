using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 수당계산식설정
    /// </summary>
    internal class PH_PY106 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.Matrix oMat03;
        private SAPbouiCOM.DBDataSource oDS_PH_PY106A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY106B;
        private SAPbouiCOM.DBDataSource oDS_PH_PY106C;
        private SAPbouiCOM.DBDataSource oDS_PH_PY106D;
        private string oLastItemUID;     //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow;         //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY106.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY106_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY106");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                oForm.Items.Item("FLD01").Specific.Select();
                PH_PY106_CreateItems();
                PH_PY106_EnableMenus();
                PH_PY106_SetDocument(oFormDocEntry);
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
        /// <returns></returns>
        private void PH_PY106_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                oDS_PH_PY106A = oForm.DataSources.DBDataSources.Item("@PH_PY106A");                ////헤더
                oDS_PH_PY106B = oForm.DataSources.DBDataSources.Item("@PH_PY106B");                ////라인
                oDS_PH_PY106C = oForm.DataSources.DBDataSources.Item("@PH_PY106C");                ////라인
                oDS_PH_PY106D = oForm.DataSources.DBDataSources.Item("@PH_PY106D");                ////라인

                oForm.DataSources.UserDataSources.Add("DISSIL", SAPbouiCOM.BoDataType.dt_LONG_TEXT);

                //공식
                oForm.Items.Item("DISSIL").Specific.DataBind.SetBound(true, "", "DISSIL");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat02 = oForm.Items.Item("Mat02").Specific;
                oMat03 = oForm.Items.Item("Mat03").Specific;

                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();
                oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat03.AutoResizeColumns();

                //----------------------------------------------------------------------------------------------
                // 헤더 설정
                //----------------------------------------------------------------------------------------------
                // 귀속년월
                oForm.Items.Item("YM").Specific.Value = DateTime.Now.ToString("yyyyMM");

                //사업장
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 1.급여형태-계산식
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P132' AND U_UseYN = 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("PAYTYP").Specific, "");
                oForm.Items.Item("PAYTYP").DisplayDesc = true;

                // 1.근속년수 계산기준
                oForm.Items.Item("GNSGBN").Specific.ValidValues.Add("1", "그룹입사일");
                oForm.Items.Item("GNSGBN").Specific.ValidValues.Add("2", "입사  일자");
                if (oForm.Items.Item("GNSGBN").Specific.ValidValues.Count >= 1)
                {
                    oForm.Items.Item("GNSGBN").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
                }

                // 2.상여 계산단위
                oForm.Items.Item("BNSLEN").Specific.ValidValues.Add("1", "  원");
                oForm.Items.Item("BNSLEN").Specific.ValidValues.Add("10", "십원");
                oForm.Items.Item("BNSLEN").Specific.ValidValues.Add("100", "백원");
                oForm.Items.Item("BNSLEN").Specific.ValidValues.Add("1000", "천원");
                if (oForm.Items.Item("BNSLEN").Specific.ValidValues.Count >= 1)
                {
                    oForm.Items.Item("BNSLEN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }

                // 3.상여 끝전처리
                oForm.Items.Item("BNSRND").Specific.ValidValues.Add("R", "반올림");
                oForm.Items.Item("BNSRND").Specific.ValidValues.Add("F", "절사");
                oForm.Items.Item("BNSRND").Specific.ValidValues.Add("C", "절상");
                if (oForm.Items.Item("BNSRND").Specific.ValidValues.Count >= 1)
                {
                    oForm.Items.Item("BNSRND").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY106_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true);                ////제거
                oForm.EnableMenu("1284", false);               ////취소
                oForm.EnableMenu("1293", true);                ////행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        private void PH_PY106_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PH_PY106_FormItemEnabled();
                    PH_PY106_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY106_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY106_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY106_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("PAYTYP").Enabled = true;
                    oForm.Items.Item("GNSGBN").Enabled = true;
                    oForm.Items.Item("BNSLEN").Enabled = true;
                    oForm.Items.Item("BNSRND").Enabled = true;

                    PH_PY106_Display_CsuItem();                    ////Mat02 초기값 가져옴

                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    //귀속년월
                    oForm.Items.Item("YM").Specific.Value = DateTime.Now.ToString("yyyyMM");

                    //1.근속년수 계산기준
                    oForm.Items.Item("GNSGBN").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);

                    //2.상여 계산단위 
                    oForm.Items.Item("BNSLEN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                    //3.상여 끝전처리
                    oForm.Items.Item("BNSRND").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                    oForm.EnableMenu("1293", true); //행삭제
                    oForm.EnableMenu("1283", true); //제거
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("PAYTYP").Enabled = true;
                    oForm.Items.Item("GNSGBN").Enabled = true;
                    oForm.Items.Item("BNSLEN").Enabled = true;
                    oForm.Items.Item("BNSRND").Enabled = true;

                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1293", true); //행삭제
                    oForm.EnableMenu("1283", true); //제거
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("YM").Enabled = false;
                    oForm.Items.Item("PAYTYP").Enabled = false;
                    oForm.Items.Item("GNSGBN").Enabled = false;
                    oForm.Items.Item("BNSLEN").Enabled = false;
                    oForm.Items.Item("BNSRND").Enabled = false;

                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);

                    oForm.EnableMenu("1293", true); //행삭제
                    oForm.EnableMenu("1283", true); //제거
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY106_DataValidCheck()
        {
            bool functionReturnValue = false;
            string sQry;
            string ExistYN = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_CLTCOD", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_YM", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("작업연월은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_PAYTYP", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("급여형태는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("PAYTYP").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_GNSGBN", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("근속일계산기준은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("GNSGBN").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_BNSLEN", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("상여 계산단위는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("BNSLEN").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_BNSRND", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("상여 끝전처리는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("BNSRND").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                //Code & Name 생성
                oDS_PH_PY106A.SetValue("Code", 0, oDS_PH_PY106A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY106A.GetValue("U_YM", 0).ToString().Trim() + oDS_PH_PY106A.GetValue("U_PAYTYP", 0).ToString().Trim());
                oDS_PH_PY106A.SetValue("NAME", 0, oDS_PH_PY106A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY106A.GetValue("U_YM", 0).ToString().Trim() + oDS_PH_PY106A.GetValue("U_PAYTYP", 0).ToString().Trim());

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    //저장된 데이터 체크
                    sQry = "SELECT Top 1 Code FROM [@PH_PY106A] ";
                    sQry += " WHERE Code = '" + oDS_PH_PY106A.GetValue("Code", 0).ToString().Trim() + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (oRecordSet01.Fields.Count > 0)
                    {
                        ExistYN = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                    }

                    if (!string.IsNullOrEmpty(ExistYN) && oDS_PH_PY106A.GetValue("Code", 0).ToString().Trim() != ExistYN)
                    {
                        PSH_Globals.SBO_Application.StatusBar.SetText("Code" + "데이터가 일치합니다. 저장되지 않습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY106_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            short ErrNumm = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY106A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    ErrNumm = 1;
                    throw new Exception();
                }

                if (ValidateType == "수정")
                {

                }
                else if (ValidateType == "행삭제")
                {

                }
                else if (ValidateType == "취소")
                {

                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (ErrNumm == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_Validate_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        private void PH_PY106_AddMatrixRow()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);

                oMat01.FlushToDataSource();
                oRow = oMat01.VisualRowCount;

                if (oMat01.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY106B.GetValue("U_CSUCOD", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY106B.Size <= oMat01.VisualRowCount)
                        {
                            oDS_PH_PY106B.InsertRecord(oRow);
                        }
                        oDS_PH_PY106B.Offset = oRow;
                        oDS_PH_PY106B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY106B.SetValue("U_LINSEQ", oRow, "");
                        oDS_PH_PY106B.SetValue("U_CSUCOD", oRow, "");
                        oDS_PH_PY106B.SetValue("U_CSUNAM", oRow, "");
                        oDS_PH_PY106B.SetValue("U_SILCUN", oRow, "");
                        oDS_PH_PY106B.SetValue("U_SILCOD", oRow, "");
                        oDS_PH_PY106B.SetValue("U_BNSBAS", oRow, "N");
                        oDS_PH_PY106B.SetValue("U_REMARK", oRow, "");
                        oMat01.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY106B.Offset = oRow - 1;
                        oDS_PH_PY106B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY106B.SetValue("U_LINSEQ", oRow - 1, "");
                        oDS_PH_PY106B.SetValue("U_CSUCOD", oRow - 1, "");
                        oDS_PH_PY106B.SetValue("U_CSUNAM", oRow - 1, "");
                        oDS_PH_PY106B.SetValue("U_SILCUN", oRow - 1, "");
                        oDS_PH_PY106B.SetValue("U_SILCOD", oRow - 1, "");
                        oDS_PH_PY106B.SetValue("U_BNSBAS", oRow - 1, "N");
                        oDS_PH_PY106B.SetValue("U_REMARK", oRow - 1, "");
                        oMat01.LoadFromDataSource();
                    }
                }
                else if (oMat01.VisualRowCount == 0)
                {
                    oDS_PH_PY106B.Offset = oRow;
                    oDS_PH_PY106B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY106B.SetValue("U_LINSEQ", oRow, "");
                    oDS_PH_PY106B.SetValue("U_CSUCOD", oRow, "");
                    oDS_PH_PY106B.SetValue("U_CSUNAM", oRow, "");
                    oDS_PH_PY106B.SetValue("U_SILCUN", oRow, "");
                    oDS_PH_PY106B.SetValue("U_SILCOD", oRow, "");
                    oDS_PH_PY106B.SetValue("U_BNSBAS", oRow, "N");
                    oDS_PH_PY106B.SetValue("U_REMARK", oRow, "");
                    oMat01.LoadFromDataSource();
                }

                //[Mat02 용]
                oMat03.FlushToDataSource();
                oRow = oMat03.VisualRowCount;

                if (oMat03.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY106D.GetValue("U_CSUCOD", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY106D.Size <= oMat03.VisualRowCount)
                        {
                            oDS_PH_PY106D.InsertRecord(oRow);
                        }
                        oDS_PH_PY106D.Offset = oRow;
                        oDS_PH_PY106D.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY106D.SetValue("U_LINSEQ", oRow, "");
                        oDS_PH_PY106D.SetValue("U_Status", oRow, "");
                        oDS_PH_PY106D.SetValue("U_WorkType", oRow, "");
                        oDS_PH_PY106D.SetValue("U_Order", oRow, "");
                        oDS_PH_PY106D.SetValue("U_CSUCOD", oRow, "");
                        oDS_PH_PY106D.SetValue("U_SILCUN", oRow, "");
                        oDS_PH_PY106D.SetValue("U_REMARK", oRow, "");
                        oMat03.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY106D.Offset = oRow - 1;
                        oDS_PH_PY106D.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY106D.SetValue("U_LINSEQ", oRow - 1, "");
                        oDS_PH_PY106D.SetValue("U_Status", oRow - 1, "");
                        oDS_PH_PY106D.SetValue("U_WorkType", oRow - 1, "");
                        oDS_PH_PY106D.SetValue("U_Order", oRow - 1, "");
                        oDS_PH_PY106D.SetValue("U_CSUCOD", oRow - 1, "");
                        oDS_PH_PY106D.SetValue("U_SILCUN", oRow - 1, "");
                        oDS_PH_PY106D.SetValue("U_REMARK", oRow - 1, "");
                        oMat03.LoadFromDataSource();
                    }
                }
                else if (oMat03.VisualRowCount == 0)
                {
                    oDS_PH_PY106D.Offset = oRow;
                    oDS_PH_PY106D.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY106D.SetValue("U_LINSEQ", oRow, "");
                    oDS_PH_PY106D.SetValue("U_Status", oRow, "");
                    oDS_PH_PY106D.SetValue("U_WorkType", oRow, "");
                    oDS_PH_PY106D.SetValue("U_Order", oRow, "");
                    oDS_PH_PY106D.SetValue("U_CSUCOD", oRow, "");
                    oDS_PH_PY106D.SetValue("U_SILCUN", oRow, "");
                    oDS_PH_PY106D.SetValue("U_REMARK", oRow, "");
                    oMat03.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY106_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY106'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY106_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스 필수 사항 check
        /// 구현은 되어 있지만 사용하지 않음
        /// </summary>
        /// <returns></returns>
        private bool MatrixSpaceLineDel()
        {
            bool functionReturnValue = false;

            int iRow;
            int kRow;
            short ErrNum = 0;
            string Chk_Data;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                oMat01.FlushToDataSource();

                if (oMat01.RowCount == 1)
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                for (iRow = 0; iRow <= oMat01.VisualRowCount - 2; iRow++)
                {
                    oDS_PH_PY106B.Offset = iRow;
                    if (string.IsNullOrEmpty(oDS_PH_PY106B.GetValue("U_CSUCOD", iRow).ToString().Trim()))
                    {
                        ErrNum = 2;
                        oMat01.Columns.Item("CSUCOD").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PH_PY106B.GetValue("U_LINSEQ", iRow).ToString().Trim()) & codeHelpClass.Left(oDS_PH_PY106B.GetValue("U_CSUCOD", iRow), 1) != "X")
                    {
                        ErrNum = 5;
                        oMat01.Columns.Item("LINSEQ").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PH_PY106B.GetValue("U_SILCUN", iRow).ToString().Trim()))
                    {
                        ErrNum = 4;
                        oMat01.Columns.Item("SILCUN").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    else if (oDS_PH_PY106B.GetValue("U_BNSBAS", iRow).ToString().Trim() == "Y")
                    {
                        if (codeHelpClass.Left(oDS_PH_PY106B.GetValue("U_CSUCOD", iRow).ToString().Trim(), 1) == "X")
                        {
                            ErrNum = 6;
                            oMat01.Columns.Item("BNSBAS").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                        else if (oDS_PH_PY106B.GetValue("U_CSUCOD", iRow).ToString().Trim() == "A04")
                        {
                            ErrNum = 7;
                            oMat01.Columns.Item("CSUCOD").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }
                    else
                    {
                        Chk_Data = oDS_PH_PY106B.GetValue("U_CSUCOD", iRow).ToString().Trim();
                        for (kRow = iRow + 1; kRow <= oMat01.VisualRowCount - 2; kRow++)
                        {
                            oDS_PH_PY106B.Offset = kRow;
                            if (Chk_Data.ToString().Trim() == oDS_PH_PY106B.GetValue("U_CSUCOD", kRow).ToString().Trim())
                            {
                                ErrNum = 3;
                                oMat01.Columns.Item("LINSEQ").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                throw new Exception();
                            }
                        }
                    }
                }

                oDS_PH_PY106B.RemoveRecord(oDS_PH_PY106B.Size - 1);
                oMat01.LoadFromDataSource();
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("입력할 데이터가 없습니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("코드는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("내용이 중복입력되었습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("계산식은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 5)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("순서는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 6)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("기본일급/통상일급/기본시급/통상시급은 상여지정을 할 수 없습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 7)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("상여금에는 상여지정을 할 수 없습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("MatrixSpaceLineDel_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }

            return functionReturnValue;
        }

        /// <summary>
        /// Display_PH_PY106
        /// </summary>
        private void Display_PH_PY106()
        {
            int i;
            int cnt;
            string sQry;
            string oCLTCOD;
            string oJOBYMM;
            string CSUCOD;
            string SILCUN;
            string SILTYP;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //Matrix2 초기화
                cnt = oDS_PH_PY106B.Size;
                if (cnt > 0)
                {
                    for (i = 1; i <= cnt - 1; i++)
                    {
                        oDS_PH_PY106B.RemoveRecord(oDS_PH_PY106B.Size - 1);
                    }
                }
                else
                {
                    oMat01.LoadFromDataSource();
                }
                oCLTCOD = oDS_PH_PY106A.GetValue("U_CLTCOD", 0).ToString().Trim();
                oJOBYMM = oDS_PH_PY106A.GetValue("U_YM", 0).ToString().Trim();
                i = 0;
                //기본셋팅값 가져오기
                sQry = "SELECT Code, Name, U_FILCOD ,U_REMARK FROM [@PH_PY106C] WHERE Code BETWEEN 'X01' AND 'X05' ORDER BY CODE";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    if (i + 1 > oDS_PH_PY106B.Size)
                    {
                        oDS_PH_PY106B.InsertRecord(i);
                    }
                    oDS_PH_PY106B.Offset = i;
                    oDS_PH_PY106B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY106B.SetValue("U_LINSEQ", i, "");
                    oDS_PH_PY106B.SetValue("U_CSUCOD", i, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                    oDS_PH_PY106B.SetValue("U_CSUNAM", i, oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oDS_PH_PY106B.SetValue("U_SILCOD", i, "");
                    oDS_PH_PY106B.SetValue("U_SILCUN", i, "");
                    oDS_PH_PY106B.SetValue("U_BNSBAS", i, "N");
                    oDS_PH_PY106B.SetValue("U_REMARK", i, oRecordSet01.Fields.Item(3).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                    i += 1;
                }
                cnt = i;
                sQry = "Exec PH_PY102 '" + oCLTCOD.ToString().Trim() + "','" + oJOBYMM.ToString().Trim() + "', '', '', '', ''";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    CSUCOD = "";
                    SILTYP = "";
                    if (i + 1 > oDS_PH_PY106B.Size)
                    {
                        oDS_PH_PY106B.InsertRecord(i);
                    }
                    CSUCOD = oRecordSet01.Fields.Item("U_CSUCOD").Value.ToString().Trim();
                    oDS_PH_PY106B.Offset = i;
                    oDS_PH_PY106B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY106B.SetValue("U_LINSEQ", i, Convert.ToString(i - cnt + 1));
                    oDS_PH_PY106B.SetValue("U_CSUCOD", i, CSUCOD);
                    oDS_PH_PY106B.SetValue("U_CSUNAM", i, oRecordSet01.Fields.Item("U_CSUNAM").Value.ToString().Trim());
                    oDS_PH_PY106B.SetValue("U_SILCOD", i, "");
                    oDS_PH_PY106B.SetValue("U_BNSBAS", i, "N");
                    oDS_PH_PY106B.SetValue("U_REMARK", i, "");
                    if (CSUCOD == "A01")
                    {
                        SILCUN = "T1.U_STDAMT";
                    }
                    else
                    {
                        SILTYP = dataHelpClass.Get_ReData("U_FIXGBN + isnull(U_INSLIN,'')", "U_CSUCOD", "[@PH_PY102B]", "'" + CSUCOD + "'", " AND Code = '" + oCLTCOD.ToString().Trim() + oJOBYMM.ToString().Trim() + "'");
                        if (codeHelpClass.Left(SILTYP, 1) == "Y")
                        {
                            SILCUN = "T2.U_CSUD" + SILTYP.Replace("Y", "").PadLeft(2, '0');
                        }
                        else
                        {
                            SILCUN = "0";
                        }
                    }

                    oDS_PH_PY106B.SetValue("U_SILCUN", i, SILCUN);

                    i += 1;
                    oRecordSet01.MoveNext();
                }
                oMat01.LoadFromDataSource();
                PH_PY106_AddMatrixRow();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Display_PH_PY106_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Display_PH_PY106
        /// </summary>
        private void PH_PY106_Display_CsuItem()
        {
            string sQry;
            int i;
            int cnt;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //Matrix2 초기화
                cnt = oDS_PH_PY106C.Size;
                if (cnt > 0)
                {
                    for (i = 1; i <= cnt - 1; i++)
                    {
                        oDS_PH_PY106C.RemoveRecord(oDS_PH_PY106C.Size - 1);
                    }
                }
                else
                {
                    oMat02.LoadFromDataSource();
                }

                i = 0;

                sQry = "SELECT Code, Name, U_FILCOD, U_REMARK FROM [@PH_PY106C] ORDER BY CODE";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    if (i + 1 > oDS_PH_PY106C.Size)
                    {
                        oDS_PH_PY106C.InsertRecord(i);
                    }
                    oDS_PH_PY106C.Offset = i;
                    oDS_PH_PY106C.SetValue("Code", i, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                    oDS_PH_PY106C.SetValue("Name", i, oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oDS_PH_PY106C.SetValue("U_FILCOD", i, oRecordSet01.Fields.Item(2).Value.ToString().Trim());
                    oDS_PH_PY106C.SetValue("U_REMARK", i, oRecordSet01.Fields.Item(3).Value.ToString().Trim());
                    i += 1;
                    oRecordSet01.MoveNext();
                }
                oMat02.LoadFromDataSource();
                oMat02.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY106_Display_CsuItem_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY106_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                if (MatrixSpaceLineDel() == false)
                                {
                                    BubbleEvent = false;
                                }
                            }
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "FLD01" || pVal.ItemUID == "FLD02")
                    {
                        oForm.PaneLevel = Convert.ToInt32(codeHelpClass.Right(pVal.ItemUID, 2));
                    }
                    if (pVal.ItemUID == "1" && pVal.ActionSuccess == true && oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PSH_Globals.SBO_Application.ActivateMenuItem("1282");
                    }
                    else if (pVal.ItemUID == "Btn1")
                    {
                        if (PH_PY106_DataValidCheck() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        else
                        {
                            Display_PH_PY106();
                        }
                    }
                    else if (pVal.ItemUID == "Btn2")
                    {
                        PH_PY106_Display_CsuItem();
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
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat1")
                        {
                            if (pVal.ColUID == "CSUCOD")
                            {
                                PH_PY106_AddMatrixRow();
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }
                        if (pVal.ItemUID == "Mat03")
                        {
                            if (pVal.ColUID == "CSUCOD")
                            {
                                PH_PY106_AddMatrixRow();
                                oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
                if (pVal.BeforeAction == true & pVal.ItemUID == "YM" & pVal.CharPressed == 9 & oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    if (oMat01.RowCount > 0)
                    {
                        oMat01.Columns.Item("LINSEQ").Cells.Item(oMat01.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        BubbleEvent = false;
                    }
                }
                else if (pVal.BeforeAction == true & pVal.ColUID == "LINSEQ" & pVal.CharPressed == 9)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
                    {
                        PSH_Globals.SBO_Application.StatusBar.SetText("순서는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false;
                    }
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
        /// Raise_EVENT_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    oMat01.LoadFromDataSource();
                    oMat02.LoadFromDataSource();
                    oMat03.LoadFromDataSource();
                    PH_PY106_FormItemEnabled();
                    PH_PY106_AddMatrixRow();
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
        /// Raise_EVENT_GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (pVal.ItemUID)
                {
                    case "Mat1":
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// Raise_EVENT_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.FormUID == oForm.UniqueID && pVal.BeforeAction == true && oLastItemUID == "Mat1" && oLastColUID == "LINSEQ" && oLastColRow > 0 && (oLastItemUID != pVal.ColUID || oLastColRow != pVal.Row) && pVal.ItemUID != "1000001" && pVal.ItemUID != "2")
                {
                    if (oLastColRow > oMat01.VisualRowCount)
                    {
                        return;
                    }
                }
                else if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "Mat1" && pVal.Row > 0)
                    {
                        oForm.DataSources.UserDataSources.Item("DISSIL").ValueEx = oMat01.Columns.Item("SILCUN").Cells.Item(pVal.Row).Specific.Value;
                    }
                    else if (pVal.ItemUID == "Mat02" && pVal.Row > 0)
                    {
                        oForm.DataSources.UserDataSources.Item("DISSIL").ValueEx = oForm.DataSources.UserDataSources.Item("DISSIL").ValueEx + oMat02.Columns.Item("FILCOD").Cells.Item(pVal.Row).Specific.Value;
                    }
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY106A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY106B);
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
		/// FORM_RESIZE 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    oForm.Items.Item("Mat02").Top = oForm.Items.Item("Mat1").Top;
                    oForm.Items.Item("Mat02").Left = oForm.Items.Item("Mat1").Width + 15;
                    oForm.Items.Item("Mat02").Width = Convert.ToInt32("240");
                    oForm.Items.Item("Mat02").Height = oForm.Items.Item("Mat1").Height;
                    oMat02.Columns.Item("Code").Width = Convert.ToInt32("20");
                    oMat02.Columns.Item("Name").Width = Convert.ToInt32("30");
                    oMat02.Columns.Item("FILCOD").Width = Convert.ToInt32("90");
                    oMat02.Columns.Item("REMARK").Width = Convert.ToInt32("80");
                    oForm.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_RESIZE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY106A", "Code"); //접속자 권한에 따른 사업장 보기
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY106_FormItemEnabled();
                            PH_PY106_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY106_FormItemEnabled();
                            PH_PY106_AddMatrixRow();
                            break;
                        case "1282": //문서추가
                            PH_PY106_FormItemEnabled();
                            PH_PY106_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY106_FormItemEnabled();
                            break;
                        case "1293": //행삭제
                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                oMat01.FlushToDataSource();

                                while ((i <= oDS_PH_PY106B.Size - 1))
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY106B.GetValue("U_LINSEQ", i)))
                                    {
                                        oDS_PH_PY106B.RemoveRecord(i);
                                        i = 0;
                                    }
                                    else
                                    {
                                        i += 1;
                                    }
                                }

                                for (i = 0; i <= oDS_PH_PY106B.Size; i++)
                                {
                                    oDS_PH_PY106B.SetValue("U_LINSEQ", i, Convert.ToString(i + 1));
                                }

                                oMat01.LoadFromDataSource();
                            }
                            PH_PY106_AddMatrixRow();
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    case "Mat1":
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}

