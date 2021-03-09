using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 대부금개별상환
    /// </summary>
    internal class PH_PY310 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY310A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY310B;
        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY310.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY310_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY310");
                
                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PH_PY310_CreateItems();
                PH_PY310_EnableMenus();
                PH_PY310_SetDocument(oFormDocEntry01);
            }
            catch(Exception ex)
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
        private void PH_PY310_CreateItems()
        {   
            try
            {
                oForm.Freeze(true);

                oDS_PH_PY310A = oForm.DataSources.DBDataSources.Item("@PH_PY310A");
                oDS_PH_PY310B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                //대출일자_S
                oForm.DataSources.UserDataSources.Add("LoanDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("LoanDate").Specific.DataBind.SetBound(true, "", "LoanDate");
                //대출일자_E

                //대부금액_S
                oForm.DataSources.UserDataSources.Add("LoanAmt", SAPbouiCOM.BoDataType.dt_SUM, 50);
                oForm.Items.Item("LoanAmt").Specific.DataBind.SetBound(true, "", "LoanAmt");
                //대부금액_E

                //총상환금액_S
                oForm.DataSources.UserDataSources.Add("TRpmtAmt", SAPbouiCOM.BoDataType.dt_SUM, 50);
                oForm.Items.Item("TRpmtAmt").Specific.DataBind.SetBound(true, "", "TRpmtAmt");
                //총상환금액_E

                //상환잔액_S
                oForm.DataSources.UserDataSources.Add("RmainAmt", SAPbouiCOM.BoDataType.dt_SUM, 50);
                oForm.Items.Item("RmainAmt").Specific.DataBind.SetBound(true, "", "RmainAmt");
                //상환잔액_E

                oMat1 = oForm.Items.Item("Mat01").Specific;

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMat1.AutoResizeColumns();

                //사업장
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
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
        /// 메뉴 세팅(Enable)
        /// </summary>
        private void PH_PY310_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false); //삭제
                oForm.EnableMenu("1287", false); //복제
                oForm.EnableMenu("1284", true); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면(Form) 초기화(Set)
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY310_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY310_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY310_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면(Form) 아이템 세팅(Enable)
        /// </summary>
        private void PH_PY310_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("LoanDoc").Enabled = true;
                    oForm.Items.Item("RpmtDate").Enabled = true;
                    oForm.Items.Item("RpmtAmt").Enabled = true;
                    oForm.Items.Item("RpmtInt").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("btnAdd").Enabled = true;

                    PH_PY310_FormClear(); //폼 DocEntry 세팅

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅

                    //대출일자,대출금액,총상환금액,상환잔액 초기화
                    oForm.Items.Item("LoanDate").Specific.Value = "";
                    oForm.Items.Item("LoanAmt").Specific.Value = "";
                    oForm.Items.Item("TRpmtAmt").Specific.Value = "";
                    oForm.Items.Item("RmainAmt").Specific.Value = "";

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가

                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("LoanDoc").Enabled = true;
                    oForm.Items.Item("RpmtDate").Enabled = true;
                    oForm.Items.Item("RpmtAmt").Enabled = true;
                    oForm.Items.Item("RpmtInt").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("btnAdd").Enabled = true;

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅

                    //대출일자,대출금액,총상환금액,상환잔액 초기화
                    oForm.Items.Item("LoanDate").Specific.Value = "";
                    oForm.Items.Item("LoanAmt").Specific.Value = "";
                    oForm.Items.Item("TRpmtAmt").Specific.Value = "";
                    oForm.Items.Item("RmainAmt").Specific.Value = "";

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가

                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("CntcCode").Enabled = false;
                    oForm.Items.Item("LoanDoc").Enabled = false;
                    oForm.Items.Item("RpmtDate").Enabled = false;
                    oForm.Items.Item("RpmtAmt").Enabled = false;
                    oForm.Items.Item("RpmtInt").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("btnAdd").Enabled = false;

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
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
        /// 화면 클리어
        /// </summary>
        private void PH_PY310_FormClear()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                string DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY310'", "");

                if (Convert.ToDouble(DocEntry) == 0)
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
        /// PH_PY310_DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY310_DataValidCheck()
        {
            bool returnValue = false;
            
            //SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //사업장
                if (string.IsNullOrEmpty(oDS_PH_PY310A.GetValue("U_CLTCOD", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //사번
                if (string.IsNullOrEmpty(oDS_PH_PY310A.GetValue("U_CntcCode", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사번은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //대부금문서번호
                if (string.IsNullOrEmpty(oDS_PH_PY310A.GetValue("U_LoanDoc", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("대부금문서는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("LoanDoc").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //상환일자
                if (string.IsNullOrEmpty(oDS_PH_PY310A.GetValue("U_RpmtDate", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("상환일자는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("RpmtDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //상환금액
                if (string.IsNullOrEmpty(oDS_PH_PY310A.GetValue("U_RpmtAmt", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("상환금액은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("RpmtAmt").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //상환이자
                if (string.IsNullOrEmpty(oDS_PH_PY310A.GetValue("U_RpmtInt", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("상환이자는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("RpmtInt").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //라인
                if (oMat1.VisualRowCount > 0)
                {
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return returnValue;
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
        /// 메트릭스에 데이터 로드
        /// </summary>
        private void PH_PY310_MTX01()
        {
            int i;
            string sQry;
            short errNum = 0;
            string Param01;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                oForm.Freeze(true);

                Param01 = oForm.Items.Item("LoanDoc").Specific.Value;

                sQry = "EXEC PH_PY310_01 '" + Param01 + "'";
                oRecordSet.DoQuery(sQry);

                oMat1.Clear();
                oMat1.FlushToDataSource();
                oMat1.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    oMat1.Clear();
                    errNum = 1;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PH_PY310B.InsertRecord(i);
                    }

                    oDS_PH_PY310B.Offset = i;
                    oDS_PH_PY310B.SetValue("U_LineNum", i, (i + 1).ToString()); //라인번호
                    oDS_PH_PY310B.SetValue("U_ColDt01", i, oRecordSet.Fields.Item("RpmtDate").Value);
                    oDS_PH_PY310B.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("RpmtAmt").Value);
                    oDS_PH_PY310B.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("RpmtInt").Value);
                    oDS_PH_PY310B.SetValue("U_ColSum03", i, oRecordSet.Fields.Item("RmainAmt").Value);
                    oDS_PH_PY310B.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("AddYN").Value);
                    oDS_PH_PY310B.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("Cnt").Value);

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
                    PSH_Globals.SBO_Application.StatusBar.SetText("기존 상환내역이 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// 대출일자, 대출금액, 총상환금액, 상환잔액 조회
        /// </summary>
        /// <param name="pLoanDoc">대부금등록 문서번호</param>
        private void PH_PY310_GetLoanData(string pLoanDoc)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "EXEC PH_PY310_02 '" + pLoanDoc + "'";
                oRecordSet.DoQuery(sQry);

                oForm.Items.Item("LoanDate").Specific.Value = oRecordSet.Fields.Item("LoanDate").Value;
                oForm.Items.Item("LoanAmt").Specific.Value = oRecordSet.Fields.Item("Loanamt").Value;
                oForm.Items.Item("TRpmtAmt").Specific.Value = oRecordSet.Fields.Item("TRpmtAmt").Value;
                oForm.Items.Item("RmainAmt").Specific.Value = oRecordSet.Fields.Item("RmainAmt").Value;
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
        /// 상환일자, 상환금액, 상환이자 입력 여부 체크
        /// </summary>
        /// <returns>성공여부</returns>
        private bool PH_PY310_CheckRepaymentData()
        {
            bool returnValue = false;
            
            try
            {
                //상환일자
                if (string.IsNullOrEmpty(oDS_PH_PY310A.GetValue("U_RpmtDate", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("상환일자는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("RpmtDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //상환금액
                if (string.IsNullOrEmpty(oDS_PH_PY310A.GetValue("U_RpmtAmt", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("상환금액은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("RpmtAmt").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //상환이자
                if (string.IsNullOrEmpty(oDS_PH_PY310A.GetValue("U_RpmtInt", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("상환이자는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("RpmtInt").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
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
        /// 개별상환내역을 Matrix에 행 추가
        /// </summary>
        private void PH_PY310_AddRepaymentData()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);

                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY310B.GetValue("U_LineNum", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY310B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY310B.InsertRecord(oRow);
                        }
                        oDS_PH_PY310B.Offset = oRow;
                        oDS_PH_PY310B.SetValue("U_LineNum", oRow, (oRow + 1).ToString());
                        oDS_PH_PY310B.SetValue("U_ColDt01", oRow, oDS_PH_PY310A.GetValue("U_RpmtDate", 0).Trim());
                        oDS_PH_PY310B.SetValue("U_ColSum01", oRow, oDS_PH_PY310A.GetValue("U_RpmtAmt", 0).Trim());
                        oDS_PH_PY310B.SetValue("U_ColSum02", oRow, oDS_PH_PY310A.GetValue("U_RpmtInt", 0).Trim());
                        oDS_PH_PY310B.SetValue("U_ColSum03", oRow, Convert.ToDouble(oForm.Items.Item("LoanAmt").Specific.Value) - Convert.ToDouble(oForm.Items.Item("TRpmtAmt").Specific.Value) - Convert.ToDouble(oDS_PH_PY310A.GetValue("U_RpmtAmt", 0).Trim()));
                        oDS_PH_PY310B.SetValue("U_ColReg01", oRow, "Y");
                        oDS_PH_PY310B.SetValue("U_ColReg02", oRow, Convert.ToString(Convert.ToDouble(oDS_PH_PY310B.GetValue("U_ColReg02", oRow - 1)) + 1)); //이전 행의 회차 + 1

                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY310B.Offset = oRow - 1;
                        oDS_PH_PY310B.SetValue("U_LineNum", oRow - 1, oRow.ToString());
                        oDS_PH_PY310B.SetValue("U_ColDt01", oRow - 1, oDS_PH_PY310A.GetValue("U_RpmtDate", 0).Trim());
                        oDS_PH_PY310B.SetValue("U_ColSum01", oRow - 1, oDS_PH_PY310A.GetValue("U_RpmtAmt", 0).Trim());
                        oDS_PH_PY310B.SetValue("U_ColSum02", oRow - 1, oDS_PH_PY310A.GetValue("U_RpmtInt", 0).Trim());
                        oDS_PH_PY310B.SetValue("U_ColSum03", oRow - 1, Convert.ToDouble(oForm.Items.Item("LoanAmt").Specific.Value) - Convert.ToDouble(oForm.Items.Item("TRpmtAmt").Specific.Value) - Convert.ToDouble(oDS_PH_PY310A.GetValue("U_RpmtAmt", 0).Trim()));
                        oDS_PH_PY310B.SetValue("U_ColReg01", oRow - 1, "Y");
                        oDS_PH_PY310B.SetValue("U_ColReg02", oRow - 1, Convert.ToString(Convert.ToDouble(oDS_PH_PY310B.GetValue("U_ColReg02", oRow - 2)) + 1)); //이전 행의 회차 + 1

                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY310B.Offset = oRow;
                    oDS_PH_PY310B.SetValue("U_LineNum", oRow, (oRow + 1).ToString());
                    oDS_PH_PY310B.SetValue("U_ColDt01", oRow, oDS_PH_PY310A.GetValue("U_RpmtDate", 0).Trim());
                    oDS_PH_PY310B.SetValue("U_ColSum01", oRow, oDS_PH_PY310A.GetValue("U_RpmtAmt", 0).Trim());
                    oDS_PH_PY310B.SetValue("U_ColSum02", oRow, oDS_PH_PY310A.GetValue("U_RpmtInt", 0).Trim());
                    oDS_PH_PY310B.SetValue("U_ColSum03", oRow, Convert.ToDouble(oForm.Items.Item("LoanAmt").Specific.Value) - Convert.ToDouble(oForm.Items.Item("TRpmtAmt").Specific.Value) - Convert.ToDouble(oDS_PH_PY310A.GetValue("U_RpmtAmt", 0).Trim()));
                    oDS_PH_PY310B.SetValue("U_ColReg01", oRow, "Y");
                    oDS_PH_PY310B.SetValue("U_ColReg02", oRow, "1"); //첫 행은 무조건 1회차

                    oMat1.LoadFromDataSource();
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
        /// 대부금개별상환 내역을 Z_PH_PY310에 INSERT, @PH_PY309B에 UPDATE
        /// </summary>
        private void PH_PY310_UpdateRepaymentData()
        {
            short loopCount;
            string sQry;

            string CLTCOD; //사업장
            string CntcCode; //사번
            short LoanDoc; //대부금문서번호
            string RpmtDate; //상환일자
            double RpmtAmt; //상환금액
            double RpmtInt; //상환이자
            double RmainAmt; //상환잔액
            short Cnt; //회차

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                for (loopCount = 0; loopCount <= oMat1.VisualRowCount - 1; loopCount++)
                {
                    if (oDS_PH_PY310B.GetValue("U_ColReg01", loopCount).Trim() == "Y")
                    {
                        CLTCOD = oDS_PH_PY310A.GetValue("U_CLTCOD", 0).Trim();
                        CntcCode = oDS_PH_PY310A.GetValue("U_CntcCode", 0).Trim();
                        LoanDoc = Convert.ToInt16(oDS_PH_PY310A.GetValue("U_LoanDoc", 0));
                        RpmtDate = oDS_PH_PY310B.GetValue("U_ColDt01", loopCount);
                        RpmtAmt = Convert.ToDouble(oDS_PH_PY310B.GetValue("U_ColSum01", loopCount));
                        RpmtInt = Convert.ToDouble(oDS_PH_PY310B.GetValue("U_ColSum02", loopCount));
                        RmainAmt = Convert.ToDouble(oDS_PH_PY310B.GetValue("U_ColSum03", loopCount));
                        Cnt = Convert.ToInt16(oDS_PH_PY310B.GetValue("U_ColReg02", loopCount));

                        sQry = "EXEC PH_PY310_03 '" + CLTCOD + "','" + CntcCode + "','" + LoanDoc + "','" + RpmtDate + "','" + RpmtAmt + "','" + RpmtInt + "','" + RmainAmt + "','" + Cnt + "'";
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY310_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            PH_PY310_UpdateRepaymentData();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY310_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "btnAdd")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY310_CheckRepaymentData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            PH_PY310_AddRepaymentData();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "btnTest")
                    {

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PH_PY310_UpdateRepaymentData();
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
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                PSH_Globals.SBO_Application.ActivateMenuItem("1291"); //이동(최종데이타)
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY310_FormItemEnabled();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY310_FormItemEnabled();
                                PH_PY310_GetLoanData(oForm.Items.Item("LoanDoc").Specific.Value); //대출일자, 대부금액, 총상환금액, 상환잔액 표시
                                PH_PY310_MTX01();
                            }
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
                    if (pVal.ItemUID == "Mat01")
                    {
                    }
                    else if (pVal.ItemUID == "CntcCode" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "LoanDoc" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("LoanDoc").Specific.Value))
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "CntcCode":

                                oDS_PH_PY310A.SetValue("U_CntcName", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", ""));
                                break;

                            case "LoanDoc":

                                if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                                {
                                    oDS_PH_PY310A.SetValue("U_CntcCode", 0, dataHelpClass.Get_ReData("U_CntcCode", "DocEntry", "[@PH_PY309A]", "'" + oForm.Items.Item("LoanDoc").Specific.Value + "'", ""));
                                    oDS_PH_PY310A.SetValue("U_CntcName", 0, dataHelpClass.Get_ReData("U_CntcName", "DocEntry", "[@PH_PY309A]", "'" + oForm.Items.Item("LoanDoc").Specific.Value + "'", ""));
                                }

                                PH_PY310_MTX01();
                                PH_PY310_GetLoanData(oForm.Items.Item("LoanDoc").Specific.Value); //대출일자, 대부금액, 총상환금액, 상환잔액 표시
                                break;

                            case "Mat01":

                                oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oMat1.AutoResizeColumns();
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
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
                    PH_PY310_FormItemEnabled();
                    oMat1.AutoResizeColumns();
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
                    SubMain.Remove_Forms(oFormUniqueID);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY310A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY310B);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
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
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
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
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY310A", "DocEntry"); //접속자 권한에 따른 사업장 보기
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY310_FormItemEnabled();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY310_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY310_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY310_FormItemEnabled();
                            PH_PY310_GetLoanData(oForm.Items.Item("LoanDoc").Specific.Value); //대출일자, 대부금액, 총상환금액, 상환잔액 표시
                            PH_PY310_MTX01();
                            break;
                        case "1293": //행삭제
                            if (oMat1.RowCount != oMat1.VisualRowCount)
                            {
                                oMat1.FlushToDataSource();

                                while (i <= oDS_PH_PY310B.Size - 1)
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY310B.GetValue("U_LineNum", i)))
                                    {
                                        oDS_PH_PY310B.RemoveRecord(i);
                                        i = 0;
                                    }
                                    else
                                    {
                                        i += 1;
                                    }
                                }

                                for (i = 0; i <= oDS_PH_PY310B.Size; i++)
                                {
                                    oDS_PH_PY310B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                }

                                oMat1.LoadFromDataSource();
                            }
                            break;
                        case "1287": //복제
                            oForm.Freeze(true);
                            oDS_PH_PY310A.SetValue("DocEntry", 0, "");

                            for (i = 0; i <= oMat1.VisualRowCount - 1; i++)
                            {
                                oMat1.FlushToDataSource();
                                oDS_PH_PY310B.SetValue("DocEntry", i, "");
                                oDS_PH_PY310B.SetValue("U_PayYN", i, "N");
                                oMat1.LoadFromDataSource();
                            }
                            oForm.Freeze(false);
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
