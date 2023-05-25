using System;
using SAPbouiCOM;
using System.Collections.Generic;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 경조사회람등록
    /// </summary>
    internal class PH_PY126 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat;
        private SAPbouiCOM.DBDataSource oDS_PH_PY126A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY126B;
        private string oLastItemUID;     //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow;         //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private bool CheckDataApply; //적용버튼 실행여부
        private string CLTCOD; //사업장
        private string YM; //적용연월
        private string DocNum; //문서번호

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY126.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY126_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY126");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";
                
                oForm.Freeze(true);
                PH_PY126_CreateItems();
                PH_PY126_ComboBox_Setting();
                PH_PY126_EnableMenus();
                PH_PY126_FormItemEnabled();
                PH_PY126_AddMatrixRow();
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
                oForm.ActiveItem = "YM"; //포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PH_PY126_CreateItems()
        {
            try
            {
                oDS_PH_PY126A = oForm.DataSources.DBDataSources.Item("@PH_PY126A");
                oDS_PH_PY126B = oForm.DataSources.DBDataSources.Item("@PH_PY126B");
                oMat = oForm.Items.Item("Mat01").Specific;
                oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMat.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// ComboBox_Setting
        /// </summary>
        private void PH_PY126_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY126_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1281", false);
                oForm.EnableMenu("1282", false);
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1285", false); // 복원
                oForm.EnableMenu("1286", false); // 닫기
                oForm.EnableMenu("1287", false); // 복제
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY126_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = true;
                    oForm.Items.Item("FieldCo").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oForm.Items.Item("Btn_Update").Enabled = true;
                    oForm.Items.Item("Btn_Cancel").Enabled = false;
                    PH_PY126_FormClear();            //폼 DocEntry 세팅

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.Items.Item("UpdateYN").Specific.Value = "N";
                    oForm.Items.Item("TAmt").Specific.Value = "0";

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.Items.Item("TAmt").Specific.Value = "0";

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("YM").Enabled = false;
                    oForm.Items.Item("CntcCode").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oForm.Items.Item("Comments").Enabled = false;

                    if (oForm.Items.Item("UpdateYN").Specific.Value.ToString().Trim() == "Y")
                    {
                        oForm.Items.Item("FieldCo").Enabled = false;
                        oForm.Items.Item("Mat01").Enabled = false;
                        oForm.Items.Item("Btn_Update").Enabled = false;
                        oForm.Items.Item("Btn_Cancel").Enabled = true;
                    }
                    else
                    {
                        oForm.Items.Item("FieldCo").Enabled = true;
                        oForm.Items.Item("Mat01").Enabled = true;
                        oForm.Items.Item("Btn_Update").Enabled = true;
                        oForm.Items.Item("Btn_Cancel").Enabled = false;
                    }

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
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
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY126_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY126'", "");
                if (Convert.ToDouble(DocEntry) == 0)
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        private void PH_PY126_AddMatrixRow()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);
                oMat.FlushToDataSource();
                oRow = oMat.VisualRowCount;

                if (oMat.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY126B.GetValue("U_LineNum", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY126B.Size <= oMat.VisualRowCount)
                        {
                            oDS_PH_PY126B.InsertRecord(oRow);
                        }
                        oDS_PH_PY126B.Offset = oRow;
                        oDS_PH_PY126B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY126B.SetValue("U_KName", oRow, "");
                        oDS_PH_PY126B.SetValue("U_KCode", oRow, "");
                        oDS_PH_PY126B.SetValue("U_KAmt", oRow, "");
                        oDS_PH_PY126B.SetValue("U_Amt", oRow, "");
                        oMat.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY126B.Offset = oRow - 1;
                        oDS_PH_PY126B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY126B.SetValue("U_KName", oRow - 1, "");
                        oDS_PH_PY126B.SetValue("U_KCode", oRow - 1, "");
                        oDS_PH_PY126B.SetValue("U_KAmt", oRow - 1, "");
                        oDS_PH_PY126B.SetValue("U_Amt", oRow - 1, "");
                        oMat.LoadFromDataSource();
                    }
                }
                else if (oMat.VisualRowCount == 0)
                {
                    oDS_PH_PY126B.Offset = oRow;
                    oDS_PH_PY126B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY126B.SetValue("U_KName", oRow, "");
                    oDS_PH_PY126B.SetValue("U_KCode", oRow, "");
                    oDS_PH_PY126B.SetValue("U_KAmt", oRow, "");
                    oDS_PH_PY126B.SetValue("U_Amt", oRow, "");
                    oMat.LoadFromDataSource();
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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY126_DataValidCheck()
        {
            bool returnValue = false;
            int i;
            int iRow;
            string errMessage = string.Empty;
            string Chk_Data;
            string Chk_Name;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oDS_PH_PY126A.GetValue("U_CLTCOD", 0).ToString().Trim()))
                {
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    errMessage = "사업장은 필수입니다.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("YM").Specific.Value.ToString().Trim()))
                {
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    errMessage = "년월은 필수입니다. 입력하세요.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
                {
                    oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    errMessage = "사번은 필수입니다. 입력하세요.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim()))
                {
                    errMessage = "경조일자는 필수입니다. 입력하세요.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("FieldCo").Specific.Value.ToString().Trim()))
                {
                    CLTCOD = oDS_PH_PY126A.GetValue("U_CLTCOD", 0).ToString().Trim();
                    YM = codeHelpClass.Right(oDS_PH_PY126A.GetValue("U_YM", 0).ToString().Trim(), 4);
                    sQry = "select U_Sequence from [@PH_PY109Z] where code ='" + CLTCOD + YM + "111'";
                    oRecordSet.DoQuery(sQry);
                    if (oRecordSet.RecordCount == 0)
                    {
                        errMessage = "급상여변동자료가 등록되지 않았습니다. 등록을 진행하세요.";
                        throw new Exception();
                    }
                    else
                    {
                        errMessage = "급상여변동자료가 동록되어 있습니다. 업데이트 필드를 선택하세요.";
                        throw new Exception();
                    }
                }
                //라인
                if (oMat.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat.VisualRowCount - 1; i++)
                    {
                        if (string.IsNullOrEmpty(oMat.Columns.Item("KName").Cells.Item(i).Specific.Value.ToString().Trim()))
                        {
                            oMat.Columns.Item("KName").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            errMessage = "셩명은 필수입니다. 입력하세요.";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oMat.Columns.Item("KCode").Cells.Item(i).Specific.Value.ToString().Trim()))
                        {
                            errMessage = "사번은 필수입니다. 입력하세요.";
                            oMat.Columns.Item("KName").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oMat.Columns.Item("KAmt").Cells.Item(i).Specific.Value.ToString().Trim()))
                        {
                            oMat.Columns.Item("KAmt").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            errMessage = "경조금액(만원)은  필수입니다. 입력하세요.";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oMat.Columns.Item("Amt").Cells.Item(i).Specific.Value.ToString().Trim()))
                        {
                            oMat.Columns.Item("Birthday").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            errMessage = "경조금액(원)울 확인하세요.";
                            throw new Exception();
                        }

                        //사번 중복 Check
                        Chk_Data = oMat.Columns.Item("KCode").Cells.Item(i).Specific.Value.ToString().Trim();
                        Chk_Name = oMat.Columns.Item("KName").Cells.Item(i).Specific.Value.ToString().Trim();
                        for (iRow = i + 1; iRow <= oMat.VisualRowCount - 1; iRow++)
                        {
                            oDS_PH_PY126B.Offset = iRow;
                            //if (Chk_Data.ToString().Trim() == oDS_PH_PY126B.GetValue("U_KCode", iRow).ToString().Trim())
                            if (Chk_Data.ToString().Trim() == oMat.Columns.Item("KCode").Cells.Item(iRow).Specific.Value.ToString().Trim())
                            {
                                errMessage = "중복된 사번이 있습니다. 사번 : " + Chk_Data + " " + Chk_Name + " 확인하세요.";
                                oMat.Columns.Item("KName").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular); //사번은 클릭할 수 없으므로 성명에..
                                throw new Exception();
                            }
                        }
                    }
                }
                else
                {
                    errMessage = "라인 데이터가 없습니다.";
                    throw new Exception();
                }

                oMat.FlushToDataSource();
                if (oDS_PH_PY126B.Size > 1)
                {
                    oDS_PH_PY126B.RemoveRecord(oDS_PH_PY126B.Size - 1);
                }
                oMat.LoadFromDataSource();

                returnValue = true;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
            return returnValue;
        }

        /// <summary>
        /// PH_PY126_DataApply
        /// </summary>
        /// <param name="CLTCOD"></param>
        /// <param name="YM"></param>
        /// <param name="DocNum"></param>
        /// <returns></returns>
        private bool PH_PY126_DataApply(string CLTCOD, string YM, string DocNum)
        {
            bool returnValue = false;
            string sQry;
            string AMTLen;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                oMat.FlushToDataSource();

                if (oForm.Items.Item("FieldCo").Specific.Value.ToString().Trim().Length == 1)
                {
                    AMTLen = Convert.ToString(Convert.ToDouble("0") + oForm.Items.Item("FieldCo").Specific.Value).ToString().Trim();
                }
                else
                {
                    AMTLen = oForm.Items.Item("FieldCo").Specific.Value.ToString().Trim();
                }

                if (PSH_Globals.SBO_Application.MessageBox("급상여 변동자료에 적용 하시겠습니까?.", 2, "Yes", "No") == 1)
                {
                    sQry = "";
                    sQry += " update [@PH_PY109B]";
                    sQry += " set U_AMT" + AMTLen + "=isnull(U_AMT" + AMTLen + ",0)  + isnull(b.U_Amt,0)";
                    sQry += " from [@PH_PY109B] a left join [@PH_PY126B] b on b.DocEntry = " + DocNum + " and a.U_MSTCOD  = b.U_KCode ";
                    sQry += " where a.code ='" + CLTCOD + codeHelpClass.Right(YM, 4) + "111'";

                    oRecordSet.DoQuery(sQry);

                    sQry = "";
                    sQry += " update [@PH_PY126A] set U_UpdateYN = 'Y' where DocEntry = " + DocNum ;

                    oRecordSet.DoQuery(sQry);

                    oForm.Items.Item("UpdateYN").Specific.Value = "Y";
                    oForm.Items.Item("CFocus").Click((SAPbouiCOM.BoCellClickType.ct_Regular)); //포커스를 인시필드로 이동(에러남)
                    oForm.Items.Item("FieldCo").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = false;
                    oForm.Items.Item("Btn_Update").Enabled = false;
                    oForm.Items.Item("Btn_Cancel").Enabled = true;

                    PSH_Globals.SBO_Application.StatusBar.SetText("급상여변동 자료에 금액이 적용 되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }

            return returnValue;
        }

        /// <summary>
        /// PH_P6_DataCancel
        /// </summary>
        /// <param name="CLTCOD"></param>
        /// <param name="YM"></param>
        /// <param name="DocNum"></param>
        /// <returns></returns>
        private bool PH_PY126_DataCancel(string CLTCOD, string YM, string DocNum)
        {
            bool returnValue = false;
            string sQry;
            string AMTLen;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                oMat.FlushToDataSource();

                if (oForm.Items.Item("FieldCo").Specific.Value.ToString().Trim().Length == 1)
                {
                    AMTLen = Convert.ToString(Convert.ToDouble("0") + oForm.Items.Item("FieldCo").Specific.Value).ToString().Trim();
                }
                else
                {
                    AMTLen = oForm.Items.Item("FieldCo").Specific.Value.ToString().Trim();
                }

                if (PSH_Globals.SBO_Application.MessageBox("급상여 변동자료에 적용(취소) 하시겠습니까?.", 2, "Yes", "No") == 1)
                {
                    sQry = "";
                    sQry += " update [@PH_PY109B]";
                    sQry += " set U_AMT" + AMTLen + "=isnull(U_AMT" + AMTLen + ",0)  - isnull(b.U_Amt,0)";
                    sQry += " from [@PH_PY109B] a left join [@PH_PY126B] b on b.DocEntry = " + DocNum + " and a.U_MSTCOD  = b.U_KCode ";
                    sQry += " where a.code ='" + CLTCOD + codeHelpClass.Right(YM, 4) + "111'";

                    oRecordSet.DoQuery(sQry);

                    sQry = "";
                    sQry += " update [@PH_PY126A] set U_UpdateYN = 'N' where DocEntry = " + DocNum;

                    oRecordSet.DoQuery(sQry);

                    oForm.Items.Item("UpdateYN").Specific.Value = "N";
                    oForm.Items.Item("FieldCo").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oForm.Items.Item("Btn_Update").Enabled = true;
                    oForm.Items.Item("Btn_Cancel").Enabled = false;

                    PSH_Globals.SBO_Application.StatusBar.SetText("급상여변동 자료에 금액이 취소적용 되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }

            return returnValue;
        }

        /// <summary>
        /// Total_Amt
        /// </summary>
        private void Total_Amt()
        {
            int tRow;
            double TAmt = 0;

            try
            {
                for (tRow = 1; tRow <= oMat.VisualRowCount; tRow++)
                {
                    oDS_PH_PY126B.Offset = tRow - 1;
                    TAmt += Convert.ToDouble(oDS_PH_PY126B.GetValue("U_Amt", tRow - 1).Replace(",", ""));
                }
                oForm.Items.Item("TAmt").Specific.Value = String.Format("{0:#,###}", TAmt);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY126_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY126_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Btn_Update")
                    {
                        CLTCOD = oDS_PH_PY126A.GetValue("U_CLTCOD", 0).ToString().Trim();
                        YM = oDS_PH_PY126A.GetValue("U_YM", 0).ToString().Trim();
                        DocNum = oDS_PH_PY126A.GetValue("DocEntry", 0).ToString().Trim();

                        if (oMat.RowCount > 1)
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                CheckDataApply = true;
                                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular); //저장
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                PH_PY126_DataApply(CLTCOD, YM, DocNum);
                            }
                        }
                        else
                        {
                            PSH_Globals.SBO_Application.MessageBox("경조금 자료가 없습니다.");
                        }
                    }
                    else if (pVal.ItemUID == "Btn_Cancel")
                    {
                        CLTCOD = oDS_PH_PY126A.GetValue("U_CLTCOD", 0).ToString().Trim();
                        YM = oDS_PH_PY126A.GetValue("U_YM", 0).ToString().Trim();
                        DocNum = oDS_PH_PY126A.GetValue("DocEntry", 0).ToString().Trim();

                        PH_PY126_DataCancel(CLTCOD, YM, DocNum);
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
                                PH_PY126_FormItemEnabled();
                                PH_PY126_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY126_FormItemEnabled();
                                PH_PY126_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY126_FormItemEnabled();
                            }
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
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "CntcCode" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "KName" && pVal.CharPressed == 9)
                        {
                            if (string.IsNullOrEmpty(oMat.Columns.Item("KName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            string errMessage = string.Empty;
            string CntcCode;
            string CntcName;
            string FieldCo;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
                        switch (pVal.ItemUID)
                        {
                            case "YM":
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    CLTCOD = oDS_PH_PY126A.GetValue("U_CLTCOD", 0).ToString().Trim();
                                    YM = codeHelpClass.Right(oDS_PH_PY126A.GetValue("U_YM", 0).ToString().Trim(), 4);

                                    if (!string.IsNullOrEmpty(oDS_PH_PY126A.GetValue("U_FieldCo", 0).ToString().Trim()))
                                    {
                                        FieldCo = " = '" + oDS_PH_PY126A.GetValue("U_FieldCo", 0).ToString().Trim();
                                    }
                                    else
                                    {
                                        FieldCo = " like '%";
                                    }

                                    sQry = "select U_Sequence from [@PH_PY109Z] where code ='" + CLTCOD + YM + "111'";
                                    oRecordSet.DoQuery(sQry);

                                    if (oRecordSet.RecordCount == 0)
                                    {
                                        errMessage = "급상여변동자료 입력은 필수입니다.";
                                        throw new Exception();
                                    }
                                    else
                                    {
                                        if (oForm.Items.Item("FieldCo").Specific.ValidValues.Count > 0)
                                        {
                                            for (int i = oForm.Items.Item("FieldCo").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                            {
                                                oForm.Items.Item("FieldCo").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                            }
                                        }
                                        sQry = "select distinct U_Sequence,U_PDName from [@PH_PY109Z] where code ='" + CLTCOD + YM + "111' order by 1";
                                        dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("FieldCo").Specific, "");
                                        oForm.Items.Item("FieldCo").DisplayDesc = true;
                                    }
                                }
                                break;

                            case "CntcCode":
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();

                                sQry = "  Select FullName = U_FullName";
                                sQry += "   From [@PH_PY001A]";
                                sQry += "  Where U_CLTCOD = '" + CLTCOD + "'";
                                sQry += "    and U_status <> '5'"; // 퇴사자 제외
                                sQry += "    and Code = '" + CntcCode + "'";
                                oRecordSet.DoQuery(sQry);

                                oForm.Items.Item("CntcName").Specific.Value = oRecordSet.Fields.Item("FullName").Value.ToString().Trim();
                                break;

                            case "Mat01":
                                if (pVal.ColUID == "KName") //성명
                                {
                                    CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                    CntcName = oMat.Columns.Item("KName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();

                                    sQry = "Select Code";
                                    sQry += " From [@PH_PY001A]";
                                    sQry += "Where U_CLTCOD = '" + CLTCOD + "'";
                                    sQry += "  And U_status <> '5'"; // 퇴사자 제외
                                    sQry += "  and U_FullName = '" + CntcName + "'";
                                    oRecordSet.DoQuery(sQry);

                                    oMat.Columns.Item("KCode").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item("Code").Value.ToString().Trim();

                                    oMat.FlushToDataSource();
                                    oDS_PH_PY126B.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oMat.LoadFromDataSource();

                                    if (oMat.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PH_PY126B.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                    {
                                        PH_PY126_AddMatrixRow();
                                    }
                                }
                                else if (pVal.ColUID == "KAmt") // 경조금액(만원)
                                {
                                    if (!string.IsNullOrEmpty(oMat.Columns.Item("KAmt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
                                    {
                                        oMat.FlushToDataSource();
                                        oDS_PH_PY126B.SetValue("U_Amt", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat.Columns.Item("KAmt").Cells.Item(pVal.Row).Specific.Value) * 10000));
                                        oMat.LoadFromDataSource();
                                    }
                                    Total_Amt();
                                }

                                oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oMat.AutoResizeColumns();
                                break;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
                oForm.Freeze(false);
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
                    oMat.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    oMat.LoadFromDataSource();

                    PH_PY126_FormItemEnabled();
                    PH_PY126_AddMatrixRow();
                    oMat.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                if (pVal.BeforeAction == true)
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY126A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY126B);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            int i = 0;

            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (oMat.RowCount != oMat.VisualRowCount)
                    {
                        oMat.FlushToDataSource();

                        while (i <= oDS_PH_PY126B.Size - 1)
                        {
                            if (string.IsNullOrEmpty(oDS_PH_PY126B.GetValue("U_LineNum", i)))
                            {
                                oDS_PH_PY126B.RemoveRecord(i);
                                i = 0;
                            }
                            else
                            {
                                i += 1;
                            }
                        }

                        for (i = 0; i <= oDS_PH_PY126B.Size; i++)
                        {
                            oDS_PH_PY126B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                        }

                        oMat.LoadFromDataSource();
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
                        case "1283":
                            if (oForm.Items.Item("UpdateYN").Specific.Value == "Y")
                            {
                                PSH_Globals.SBO_Application.MessageBox("급상여변동자료에 반영된 자료는 제거할 수 없습니다.");
                                BubbleEvent = false;
                                return;
                            }
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
                            if (oForm.Items.Item("UpdateYN").Specific.Value == "Y")
                            {
                                PSH_Globals.SBO_Application.MessageBox("급상여변동자료에 반영된 자료는 행삭제할 수 없습니다.");
                                BubbleEvent = false;
                                return;
                            }
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
                            PH_PY126_FormItemEnabled();
                            PH_PY126_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY126_FormItemEnabled();
                            PH_PY126_AddMatrixRow();
                            break;
                        case "1282": //문서추가
                            PH_PY126_FormItemEnabled();
                            PH_PY126_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY126_FormItemEnabled();
                            Total_Amt();
                            break;
                        case "1293": // 행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            // 마지막행이 공백이 아니면 추가(마지막행을 지울 경우)
                            if (!string.IsNullOrEmpty(oDS_PH_PY126B.GetValue("U_KName", oDS_PH_PY126B.Size - 1).ToString().Trim()))
                            {
                                PH_PY126_AddMatrixRow();
                                PSH_Globals.SBO_Application.StatusBar.SetText("마지막행(공백)은 행삭제하면 안됩니다..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            }
                            Total_Amt();
                            break;
                        case "1287": //복제
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }
    }
}

