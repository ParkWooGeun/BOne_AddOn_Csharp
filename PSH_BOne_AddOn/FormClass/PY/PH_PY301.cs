using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 학자금신청등록
    /// </summary>
    internal class PH_PY301 : PSH_BaseClass
    {
        public string oFormUniqueID;
        //public SAPbouiCOM.Form oForm;

        public SAPbouiCOM.Matrix oMat1;

        private SAPbouiCOM.DBDataSource oDS_PH_PY301A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY301B;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY301.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY301_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY301");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                //oForm.Visible = true;
                PH_PY301_CreateItems();
                PH_PY301_EnableMenus();
                PH_PY301_SetDocument(oFromDocEntry01);
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
        private void PH_PY301_CreateItems()
        {
            string sQry = string.Empty;
           
            SAPbobsCOM.Recordset oRecordSet = null;
            //SAPbouiCOM.ComboBox oCombo = null;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            oForm.Freeze(true);

            try
            {
                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oDS_PH_PY301A = oForm.DataSources.DBDataSources.Item("@PH_PY301A");
                oDS_PH_PY301B = oForm.DataSources.DBDataSources.Item("@PH_PY301B");

                oMat1 = oForm.Items.Item("Mat01").Specific;

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMat1.AutoResizeColumns();

                //사업장
                //oCombo = oForm.Items.Item("CLTCOD").Specific;
                //oForm.Items.Item("CLTCOD").DisplayDesc = true;


                //oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                //oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //분기
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("", "");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("01", "1/4 혹은 1학기");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("02", "2/4");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("03", "3/4 혹은 2학기");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("04", "4/4");
                oForm.Items.Item("Quarter").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Quarter").DisplayDesc = true;

                //매트릭스-성별
                oMat1.Columns.Item("Sex").ValidValues.Add("", "");
                oMat1.Columns.Item("Sex").ValidValues.Add("01", "남자");
                oMat1.Columns.Item("Sex").ValidValues.Add("02", "여자");
                oMat1.Columns.Item("Sex").DisplayDesc = true;

                //매트릭스-학교
                oMat1.Columns.Item("SchCls").ValidValues.Add("", "");
                sQry = "       SELECT   T1.U_Code,";
                sQry = sQry + "         T1.U_CodeNm";
                sQry = sQry + "  FROM   [@PS_HR200H] AS T0";
                sQry = sQry + "         INNER JOIN";
                sQry = sQry + "         [@PS_HR200L] AS T1";
                sQry = sQry + "         ON T0.Code = T1.Code";
                sQry = sQry + "  WHERE  T0.Code = 'P222'";
                sQry = sQry + "    AND T1.U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY  T1.U_Seq";

                dataHelpClass.GP_MatrixSetMatComboList(oMat1.Columns.Item("SchCls"), sQry, "", "");

                //매트릭스-학년
                oMat1.Columns.Item("Grade").ValidValues.Add("", "");
                oMat1.Columns.Item("Grade").ValidValues.Add("01", "1학년");
                oMat1.Columns.Item("Grade").ValidValues.Add("02", "2학년");
                oMat1.Columns.Item("Grade").ValidValues.Add("03", "3학년");
                oMat1.Columns.Item("Grade").ValidValues.Add("04", "4학년");
                oMat1.Columns.Item("Grade").ValidValues.Add("05", "5학년");
                oMat1.Columns.Item("Grade").DisplayDesc = true;

                //매트릭스-회차
                oMat1.Columns.Item("Count").ValidValues.Add("", "");
                oMat1.Columns.Item("Count").ValidValues.Add("01", "1차");
                oMat1.Columns.Item("Count").ValidValues.Add("02", "2차");
                oMat1.Columns.Item("Count").DisplayDesc = true;

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);  //메모리 해제
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY301_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false); // 삭제
                oForm.EnableMenu("1287", true); // 복제
                //oForm.EnableMenu("1286", True); // 닫기
                oForm.EnableMenu("1284", true); // 취소
                oForm.EnableMenu("1293", true); // 행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY301_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PH_PY301_FormItemEnabled();
                    PH_PY301_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY301_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY301_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY301_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            //SAPbouiCOM.ComboBox oCombo = null;

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("StdYear").Enabled = true;
                    oForm.Items.Item("Quarter").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;

                    //폼 DocEntry 세팅
                    PH_PY301_FormClear();

                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    //년도 세팅
                    oDS_PH_PY301A.SetValue("U_StdYear", 0, DateTime.Now.ToString("yyyy"));

                    oForm.EnableMenu("1281", true);
                    ////문서찾기
                    oForm.EnableMenu("1282", false);
                    ////문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("StdYear").Enabled = true;
                    oForm.Items.Item("Quarter").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;

                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("CntcCode").Enabled = false;
                    oForm.Items.Item("StdYear").Enabled = false;
                    oForm.Items.Item("Quarter").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;

                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);

                    oForm.EnableMenu("1281", true);
                    ////문서찾기
                    oForm.EnableMenu("1282", true);
                    ////문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Matirx 행 추가
        /// </summary>
        private void PH_PY301_AddMatrixRow()
        {
            int oRow = 0;

            try
            {
                oForm.Freeze(true);
                //[Mat1]
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;
                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY301B.GetValue("U_Name", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY301B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY301B.InsertRecord((oRow));
                        }
                        oDS_PH_PY301B.Offset = oRow;
                        oDS_PH_PY301B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY301B.SetValue("U_Name", oRow, "");
                        oDS_PH_PY301B.SetValue("U_GovID", oRow, "");
                        oDS_PH_PY301B.SetValue("U_Sex", oRow, "");
                        oDS_PH_PY301B.SetValue("U_SchCls", oRow, "");
                        oDS_PH_PY301B.SetValue("U_SchName", oRow, "");
                        oDS_PH_PY301B.SetValue("U_Grade", oRow, "");
                        oDS_PH_PY301B.SetValue("U_EntFee", oRow, Convert.ToString(0));
                        oDS_PH_PY301B.SetValue("U_Tuition", oRow, Convert.ToString(0));
                        oDS_PH_PY301B.SetValue("U_Count", oRow, "");
                        oDS_PH_PY301B.SetValue("U_PayCnt", oRow, "");
                        oDS_PH_PY301B.SetValue("U_PayYN", oRow, "");
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY301B.Offset = oRow - 1;
                        oDS_PH_PY301B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY301B.SetValue("U_Name", oRow - 1, "");
                        oDS_PH_PY301B.SetValue("U_GovID", oRow - 1, "");
                        oDS_PH_PY301B.SetValue("U_Sex", oRow - 1, "");
                        oDS_PH_PY301B.SetValue("U_SchCls", oRow - 1, "");
                        oDS_PH_PY301B.SetValue("U_SchName", oRow - 1, "");
                        oDS_PH_PY301B.SetValue("U_Grade", oRow - 1, "");
                        oDS_PH_PY301B.SetValue("U_EntFee", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY301B.SetValue("U_Tuition", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY301B.SetValue("U_Count", oRow - 1, "");
                        oDS_PH_PY301B.SetValue("U_PayCnt", oRow, "");
                        oDS_PH_PY301B.SetValue("U_PayYN", oRow - 1, "");
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY301B.Offset = oRow;
                    oDS_PH_PY301B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY301B.SetValue("U_Name", oRow, "");
                    oDS_PH_PY301B.SetValue("U_GovID", oRow, "");
                    oDS_PH_PY301B.SetValue("U_Sex", oRow, "");
                    oDS_PH_PY301B.SetValue("U_SchCls", oRow, "");
                    oDS_PH_PY301B.SetValue("U_SchName", oRow, "");
                    oDS_PH_PY301B.SetValue("U_Grade", oRow, "");
                    oDS_PH_PY301B.SetValue("U_EntFee", oRow, Convert.ToString(0));
                    oDS_PH_PY301B.SetValue("U_Tuition", oRow, Convert.ToString(0));
                    oDS_PH_PY301B.SetValue("U_Count", oRow, "");
                    oDS_PH_PY301B.SetValue("U_PayCnt", oRow, "");
                    oDS_PH_PY301B.SetValue("U_PayYN", oRow, "");
                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY301_AddMatrixRow_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY301_FormClear()
        {
            string DocEntry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY301'", "");
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY301_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY301_DataValidCheck()
        {
            bool functionReturnValue = false;

            functionReturnValue = false;

            int i = 0;
            string CLTCOD = string.Empty;
            string StdYear = string.Empty;
            string Quarter = string.Empty;
            string Count = string.Empty;
            string sQry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //사업장
                if (string.IsNullOrEmpty(oDS_PH_PY301A.GetValue("U_CLTCOD", 0)))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }
                //년도
                if (string.IsNullOrEmpty(oDS_PH_PY301A.GetValue("U_StdYear", 0)))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("년도는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("StdYear").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }
                //사번
                if (string.IsNullOrEmpty(oDS_PH_PY301A.GetValue("U_CntcCode", 0)))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사번은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }
                //분기
                if (string.IsNullOrEmpty(oDS_PH_PY301A.GetValue("U_Quarter", 0)))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("분기는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Quarter").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }
                CLTCOD = oDS_PH_PY301A.GetValue("U_CLTCOD", 0);
                StdYear = oDS_PH_PY301A.GetValue("U_StdYear", 0);
                Quarter = oDS_PH_PY301A.GetValue("U_Quarter", 0);

                //라인
                if (oMat1.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                    {

                        //학교
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("SchCls").Cells.Item(i).Specific.VALUE))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("학교는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("SchCls").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            functionReturnValue = false;
                            return functionReturnValue;
                        }
                        //학교명
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("SchName").Cells.Item(i).Specific.VALUE))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("학교명은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("SchName").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            functionReturnValue = false;
                            return functionReturnValue;
                        }
                        //학년
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("Grade").Cells.Item(i).Specific.VALUE))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("학년은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("Grade").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            functionReturnValue = false;
                            return functionReturnValue;
                        }
                        //회차
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("Count").Cells.Item(i).Specific.VALUE))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("회차는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("Count").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            functionReturnValue = false;
                            return functionReturnValue;
                        }
                        Count = oMat1.Columns.Item("Count").Cells.Item(i).Specific.VALUE;

                        sQry = "Select Cnt = Count(*) From [@PH_PY301A] a Inner Join [@PH_PY301B] b On a.DocEntry = b.DocEntry and a.Canceled = 'N' ";
                        sQry = sQry + " Where a.U_CLTCOD = '" + CLTCOD + "' And a.U_StdYear = '" + StdYear + "' and a.U_Quarter = '" + Quarter + "' ";
                        sQry = sQry + " And b.U_Count = '" + Count + "' and b.U_PayYN = 'Y'";

                        oRecordSet.DoQuery(sQry);

                        if (oRecordSet.Fields.Item(0).Value > 0)
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("지급완료처리가 되어 추가/수정을 할 수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            functionReturnValue = false;
                            return functionReturnValue;
                        }
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                oMat1.FlushToDataSource();
                //// Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY301B.Size > 1)
                    oDS_PH_PY301B.RemoveRecord((oDS_PH_PY301B.Size - 1));

                oMat1.LoadFromDataSource();

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return functionReturnValue;
        }

        //    return functionReturnValue;
        //}

        ///// <summary>
        ///// 대상자 조회
        ///// </summary>
        //private void PH_PY301_MTX01()
        //{
        //    int i = 0;
        //    string sQry = string.Empty;

        //    string YM = string.Empty;
        //    string DocDate = string.Empty;

        //    string Param01 = string.Empty;
        //    string Param02 = string.Empty;
        //    string Param03 = string.Empty;
        //    double Param04 = 0;
        //    double Param05 = 0;
        //    string Param06 = string.Empty;

        //    short ErrNum = 0;

        //    SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", 100, false);

        //    //PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

        //    try
        //    {
        //        oForm.Freeze(true);

        //        Param01 = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
        //        Param02 = oForm.Items.Item("JIGTYP").Specific.Value.ToString().Trim();
        //        Param03 = oForm.Items.Item("Gubun").Specific.Value.ToString().Trim();
        //        Param04 = Convert.ToDouble(oForm.Items.Item("UPHOBONG").Specific.Value);
        //        Param05 = Convert.ToDouble(oForm.Items.Item("UPAMT").Specific.Value);
        //        Param06 = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

        //        YM = oForm.Items.Item("YM").Specific.VALUE;
        //        DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

        //        if (string.IsNullOrEmpty(Param03))
        //        {
        //            ErrNum = 1;
        //            throw new Exception();
        //        }

        //        if (string.IsNullOrEmpty(Param06))
        //        {
        //            ErrNum = 2;
        //            throw new Exception();
        //        }

        //        sQry = "Select Count(*) From [@PH_PY301A] Where U_CLTCOD = '" + Param01 + "' and U_YM = '" + YM + "' and U_JIGTYP = '" + Param02 + "'";
        //        sQry = sQry + " and U_DocDate = '" + DocDate + "'";
        //        sQry = sQry + " and Canceled = 'N' ";
        //        oRecordSet.DoQuery(sQry);

        //        if (oRecordSet.Fields.Item(0).Value > 0)
        //        {
        //            ErrNum = 3;
        //            throw new Exception();
        //        }

        //        sQry = "EXEC [PH_PY301_01] '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', " + Param04 + ", " + Param05 + ", '" + Param06 + "'";
        //        oRecordSet.DoQuery(sQry);

        //        oMat1.Clear();
        //        oMat1.FlushToDataSource();
        //        oMat1.LoadFromDataSource();

        //        if (oRecordSet.RecordCount == 0)
        //        {
        //            ErrNum = 4;
        //            throw new Exception();
        //        }

        //        for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
        //        {
        //            if (i != 0)
        //            {
        //                oDS_PH_PY301B.InsertRecord((i));
        //            }
        //            oDS_PH_PY301B.Offset = i;
        //            oDS_PH_PY301B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
        //            oDS_PH_PY301B.SetValue("U_TeamCode", i, oRecordSet.Fields.Item(0).Value);
        //            oDS_PH_PY301B.SetValue("U_RspCode", i, oRecordSet.Fields.Item(1).Value);
        //            oDS_PH_PY301B.SetValue("U_ClsCode", i, oRecordSet.Fields.Item(2).Value);
        //            oDS_PH_PY301B.SetValue("U_MSTCOD", i, oRecordSet.Fields.Item(3).Value);
        //            oDS_PH_PY301B.SetValue("U_FULLNAME", i, oRecordSet.Fields.Item(4).Value);
        //            oDS_PH_PY301B.SetValue("U_JIGCOD", i, oRecordSet.Fields.Item(5).Value);
        //            oDS_PH_PY301B.SetValue("U_GrpDat", i, oRecordSet.Fields.Item(6).Value);
        //            oDS_PH_PY301B.SetValue("U_birthDat", i, oRecordSet.Fields.Item(7).Value);
        //            oDS_PH_PY301B.SetValue("U_HOBYMM", i, oRecordSet.Fields.Item(8).Value);
        //            oDS_PH_PY301B.SetValue("U_HOBONG", i, oRecordSet.Fields.Item(9).Value);
        //            oDS_PH_PY301B.SetValue("U_HOBNAM", i, oRecordSet.Fields.Item(10).Value);
        //            oDS_PH_PY301B.SetValue("U_STDAMT", i, oRecordSet.Fields.Item(11).Value);
        //            oDS_PH_PY301B.SetValue("U_BNSAMT", i, oRecordSet.Fields.Item(12).Value);
        //            oDS_PH_PY301B.SetValue("U_UPHOBONG", i, oRecordSet.Fields.Item(13).Value);
        //            oDS_PH_PY301B.SetValue("U_UPHOBNAM", i, oRecordSet.Fields.Item(14).Value);
        //            oDS_PH_PY301B.SetValue("U_UPSTDAMT", i, oRecordSet.Fields.Item(15).Value);
        //            oDS_PH_PY301B.SetValue("U_UPBNSAMT", i, oRecordSet.Fields.Item(16).Value);
        //            oDS_PH_PY301B.SetValue("U_APPLYYN", i, oRecordSet.Fields.Item(17).Value);
        //            oDS_PH_PY301B.SetValue("U_PeakYN", i, oRecordSet.Fields.Item(18).Value);
        //            oDS_PH_PY301B.SetValue("U_LineMemo", i, oRecordSet.Fields.Item(19).Value);
        //            oRecordSet.MoveNext();
        //            ProgressBar01.Value = ProgressBar01.Value + 1;
        //            ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
        //        }
        //        oMat1.LoadFromDataSource();
        //        oMat1.AutoResizeColumns();
        //        oForm.Update();
        //    }
        //    catch (Exception ex)
        //    {
        //        ProgressBar01.Stop(); //StatusBar를 ProgressBar01가 점유하고 있기 때문에 오류 메시지를 출력하기 위해 ProgressBar01 정지

        //        if (ErrNum == 1)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("승호기준은 필수입니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //            //dataHelpClass.MDC_GF_Message("승호기준은 필수입니다. 확인바랍니다.", "E");
        //            oForm.Items.Item("Gubun").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //        }
        //        else if (ErrNum == 2)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("승호일자는 필수입니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //            //dataHelpClass.MDC_GF_Message("승호일자는 필수입니다. 확인바랍니다.", "E");
        //            oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //        }
        //        else if (ErrNum == 3)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("승호작업을 이미 등록했습니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //            //dataHelpClass.MDC_GF_Message("승호작업을 이미 등록했습니다. 확인바랍니다.", "E");
        //        }
        //        else if (ErrNum == 4)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //            //dataHelpClass.MDC_GF_Message("결과가 존재하지 않습니다.", "E");
        //        }
        //        else
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY301_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //        }
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //        if (ProgressBar01 != null)
        //        {
        //            ProgressBar01.Stop();
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
        //        }
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
        //    }
        //}

        ///// <summary>
        ///// 승호처리 및 발령사항 추가
        ///// </summary>
        //private void PH_PY301_MTX02()
        //{

        //    //int i = 0;
        //    string sQry = string.Empty;

        //    string Param01 = string.Empty;
        //    string Param02 = string.Empty;
        //    string Param03 = string.Empty;

        //    string CLTCOD = string.Empty;
        //    string DocDate = string.Empty;
        //    string appNum = string.Empty;
        //    string JIGTYP = string.Empty;

        //    short ErrNum = 0;

        //    SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    //PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

        //    try
        //    {
        //        oForm.Freeze(true);

        //        Param01 = oForm.Items.Item("DocEntry").Specific.Value;
        //        Param02 = oForm.Items.Item("appNum").Specific.Value;
        //        Param03 = oForm.Items.Item("Canceled").Specific.Value;

        //        CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
        //        DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
        //        appNum = oForm.Items.Item("appNum").Specific.Value;
        //        JIGTYP = oForm.Items.Item("JIGTYP").Specific.Value;

        //        sQry = "Select Count(*) From [@PH_PY301A] a Inner Join [@PH_PY301B] b On a.DocEntry = b.DocEntry ";
        //        sQry = sQry + " Inner Join [@PH_PY001G] c On b.U_MSTCOD = c.Code ";
        //        sQry = sQry + " Where a.DocEntry = '" + Param01 + "' and c.U_appNum = a.U_appNum and c.U_appType = '08' ";
        //        sQry = sQry + " and c.U_appDate = '" + DocDate + "'";

        //        oRecordSet.DoQuery(sQry);
        //        if (oRecordSet.Fields.Item(0).Value > 0)
        //        {
        //            ErrNum = 1;
        //            throw new Exception();
        //        }

        //        sQry = "EXEC [PH_PY301_02] '" + Param01 + "', '" + Param02 + "', '" + Param03 + "'";
        //        oRecordSet.DoQuery(sQry);

        //        sQry = "EXEC [PH_PY301_03] '" + Param01 + "', '" + Param03 + "'";
        //        oRecordSet.DoQuery(sQry);

        //        PSH_Globals.SBO_Application.StatusBar.SetText("승호작업을 정상 처리했습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
        //        //dataHelpClass.MDC_GF_Message("승호작업을 정상 처리했습니다.", "S");
        //    }
        //    catch (Exception ex)
        //    {
        //        if (ErrNum == 1)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("승호작업을 이미 등록했습니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //            //dataHelpClass.MDC_GF_Message("승호작업을 이미 처리했습니다. 확인바랍니다.", "E");
        //        }
        //        else
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY301_MTX02_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //        }
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
        //    }
        //}

        ///// <summary>
        ///// 승호취소 및 발령관리 삭제
        ///// </summary>
        //private void PH_PY301_MTX03()
        //{
        //    //int i = 0;
        //    string sQry = null;
        //    string DocDate = null;

        //    string Param01 = null;
        //    string Param02 = null;
        //    string Param03 = null;

        //    short ErrNum = 0;

        //    SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    //PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

        //    try
        //    {
        //        oForm.Freeze(true);

        //        Param01 = oForm.Items.Item("DocEntry").Specific.Value;
        //        Param02 = oForm.Items.Item("appNum").Specific.Value;
        //        Param03 = oForm.Items.Item("Canceled").Specific.Value;

        //        DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

        //        sQry = "Select Count(*) From [@PH_PY301A] a Inner Join [@PH_PY301B] b On a.DocEntry = b.DocEntry ";
        //        sQry = sQry + " Inner Join [@PH_PY001G] c On b.U_MSTCOD = c.Code ";
        //        sQry = sQry + " Where a.DocEntry = '" + Param01 + "' and c.U_appNum = a.U_appNum and c.U_appType = '08' ";
        //        sQry = sQry + " and c.U_appDate = '" + DocDate + "'";

        //        oRecordSet.DoQuery(sQry);
        //        if (oRecordSet.Fields.Item(0).Value <= 0)
        //        {
        //            ErrNum = 1;
        //            throw new Exception();
        //        }

        //        sQry = "EXEC [PH_PY301_02] '" + Param01 + "', '" + Param02 + "', '" + Param03 + "'";
        //        oRecordSet.DoQuery(sQry);

        //        sQry = "EXEC [PH_PY301_03] '" + Param01 + "', '" + Param03 + "'";
        //        oRecordSet.DoQuery(sQry);

        //        PSH_Globals.SBO_Application.StatusBar.SetText("승호작업을 취소 처리했습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
        //        //dataHelpClass.MDC_GF_Message("승호작업을 취소 처리했습니다.", "S");
        //    }
        //    catch (Exception ex)
        //    {
        //        if (ErrNum == 1)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("승호취소할 자료가 대상이 없거나 이미 취소했습니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //            //dataHelpClass.MDC_GF_Message("승호취소 할 자료가 대상이 없거나 이미 취소했습니다. 확인바랍니다.", "E");
        //        }
        //        else
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY301_MTX03_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //        }
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
        //    }
        //}

        ///// <summary>
        ///// Validate
        ///// </summary>
        ///// <param name="ValidateType"></param>
        ///// <returns></returns>
        //private bool PH_PY301_Validate(string ValidateType)
        //{
        //    bool functionReturnValue = false;

        //    PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
        //    SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    try
        //    {
        //        if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY301A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y")
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //            return functionReturnValue;
        //        }

        //        if (ValidateType == "수정")
        //        {

        //        }
        //        else if (ValidateType == "행삭제")
        //        {

        //        }
        //        else if (ValidateType == "취소")
        //        {

        //        }

        //        functionReturnValue = true;
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY301_Validate_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }

        //    return functionReturnValue;
        //}

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
                    ///Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    ///Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
                                                             /// Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                //    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    ///Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                    //case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                    //    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
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
                            if (PH_PY301_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }

                            ////해야할일 작업
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY301_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                            ////해야할일 작업

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
                                PH_PY301_FormItemEnabled();
                                PH_PY301_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY301_FormItemEnabled();
                                PH_PY301_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY301_FormItemEnabled();
                            }
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
        /// ITEM_PRESSED 이벤트
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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string GovID = string.Empty;
            string GovID1 = string.Empty;
            string SchCls = string.Empty;
            string Sex = string.Empty;
            string GovID2 = string.Empty;

            short loopCount = 0;
            double FeeTot = 0;
            double TuiTot = 0;
            double Total = 0;
            double PreTuition = 0;
            double Tuition = 0;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "CntcCode":
                                oDS_PH_PY301A.SetValue("U_CntcName", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.VALUE + "'",""));
                                break;

                            case "Mat01":

                                if (pVal.ColUID == "Name")
                                {

                                    oMat1.FlushToDataSource();

                                    GovID = dataHelpClass.Get_ReData( "U_FamPer",  "U_FamNam",  "[@PH_PY001D]",  "'" + oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE + "'",  " AND Code = '" + oDS_PH_PY301A.GetValue("U_CntcCode", 0) + "'");

                                    if (GovID.Substring(6,1) == "1" | GovID.Substring(6, 1) == "3" | GovID.Substring(6, 1) == "5")
                                    {
                                        Sex = "01";
                                    }
                                    else if (GovID.Substring(6, 1) == "2" | GovID.Substring(6, 1) == "4" | GovID.Substring(6, 1) == "6")
                                    {
                                        Sex = "02";
                                    }

                                    GovID1 = GovID.Substring(0, 6);
                                    GovID2 = GovID.Substring(6, 7);
                                    GovID = GovID1 + "-" + GovID2;

                                    oDS_PH_PY301B.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE);
                                    oDS_PH_PY301B.SetValue("U_GovID", pVal.Row - 1, GovID);
                                    //주민등록번호
                                    oDS_PH_PY301B.SetValue("U_Sex", pVal.Row - 1, Sex);
                                    //성별
                                    oDS_PH_PY301B.SetValue("U_PayYN", pVal.Row - 1, "N");
                                    //지급완료여부

                                    oMat1.LoadFromDataSource();

                                    PH_PY301_AddMatrixRow();

                                    //입학금 입력 시
                                }
                                else if (pVal.ColUID == "EntFee")
                                {
                                    oMat1.FlushToDataSource();

                                    //학교선택을 하지 않으면 에러 메시지 출력
                                    if (string.IsNullOrEmpty(oDS_PH_PY301B.GetValue("U_SchCls", pVal.Row - 1)))
                                    {
                                        dataHelpClass.MDC_GF_Message("학교를 먼저 선택하십시오.", "E");
                                        oDS_PH_PY301B.SetValue("U_EntFee", pVal.Row - 1, Convert.ToString(0));
                                        BubbleEvent = false;
                                    }

                                    //입학금 합계 계산
                                    for (loopCount = 0; loopCount <= oMat1.RowCount - 1; loopCount++)
                                    {
                                        FeeTot = FeeTot + Convert.ToDouble(oDS_PH_PY301B.GetValue("U_EntFee", loopCount));
                                    }
                                    oMat1.LoadFromDataSource();

                                    oDS_PH_PY301A.SetValue("U_FeeTot", 0, Convert.ToString(FeeTot));

                                    TuiTot = Convert.ToDouble(oDS_PH_PY301A.GetValue("U_TuiTot", 0));
                                    Total = FeeTot + TuiTot;

                                    oDS_PH_PY301A.SetValue("U_Total", 0, Convert.ToString(Total));
                                    
                                    //등록금 입력 시
                                }
                                else if (pVal.ColUID == "Tuition")
                                {

                                    PreTuition = Convert.ToDouble(oDS_PH_PY301B.GetValue("U_Tuition", pVal.Row - 1));

                                    oMat1.FlushToDataSource();

                                    //학교선택을 하지 않으면 에러 메시지 출력
                                    if (string.IsNullOrEmpty(oDS_PH_PY301B.GetValue("U_SchCls", pVal.Row - 1)))
                                    {
                                        dataHelpClass.MDC_GF_Message( "학교를 먼저 선택하십시오.",  "E");
                                        oDS_PH_PY301B.SetValue("U_Tuition", pVal.Row - 1, Convert.ToString(0));
                                        BubbleEvent = false;
                                    }

                                    //한도금액 체크
                                    SchCls = oDS_PH_PY301B.GetValue("U_SchCls", pVal.Row - 1);
                                    Tuition = Convert.ToDouble(oDS_PH_PY301B.GetValue("U_Tuition", pVal.Row - 1));

                                    //고등학교 이외만 체크
                                    if (Convert.ToInt16(SchCls) > 2)
                                    {
                                        //if (PH_PY301_CheckAmt(Tuition, SchCls) == false)
                                        //{
                                        //    dataHelpClass.MDC_GF_Message("등록금이 한도금액을 초과하였습니다. 확인하십시오.", "E");
                                        //    oDS_PH_PY301B.SetValue("U_Tuition", pVal.Row - 1, Convert.ToString(PreTuition));
                                        //    //이전 데이터로 회귀
                                        //    BubbleEvent = false;
                                        //}
                                    }

                                    //등록금 합계 계산
                                    for (loopCount = 0; loopCount <= oMat1.RowCount - 1; loopCount++)
                                    {
                                        TuiTot = TuiTot + Convert.ToDouble(oDS_PH_PY301B.GetValue("U_Tuition", loopCount));
                                    }
                                    oMat1.LoadFromDataSource();

                                    oDS_PH_PY301A.SetValue("U_TuiTot", 0, Convert.ToString(TuiTot));

                                    FeeTot = Convert.ToDouble(oDS_PH_PY301A.GetValue("U_FeeTot", 0));
                                    Total = FeeTot + TuiTot;

                                    oDS_PH_PY301A.SetValue("U_Total", 0, Convert.ToString(Total));

                                }
                                oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oMat1.AutoResizeColumns();
                                break;
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
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "Name" & pVal.CharPressed == Convert.ToDouble("9"))
                        {
                            //UPGRADE_WARNING: oMat1.Columns.Item(Name).Cells(pVal.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            if (string.IsNullOrEmpty(oMat1.Columns.Item("Name").Cells.Item(pVal.Row).Specific.VALUE))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                    }
                    else if (pVal.ItemUID == "CntcCode" & pVal.CharPressed == Convert.ToDouble("9"))
                    {

                        //UPGRADE_WARNING: oForm.Items(CntcCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.VALUE))
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_KEY_DOWN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        ///// <summary>
        /////  등록금(학자금)의 한도금액 체크
        ///// </summary>
        ///// <param name="pAmt"></param>
        ///// <param name="pSchCls"></param>
        ///// <returns></returns>
        //private bool PH_PY301_CheckAmt(double pAmt, string pSchCls)
        //{
        //    bool functionReturnValue = false;
        //    string sQry = string.Empty;
        //    double CheckAmt = 0;

        //    SAPbobsCOM.Recordset oRecordSet = null;
        //    oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    try
        //    {
        //        sQry = "      SELECT      U_Num1 AS [CheckAmt]";
        //        sQry = sQry + " FROM      [@PS_HR200L] AS T0 ";
        //        sQry = sQry + "WHERE      T0.Code = 'P222'";
        //        sQry = sQry + "  AND T0.U_Code = '" + pSchCls.Trim() + "'";

        //        oRecordSet.DoQuery(sQry);

        //        CheckAmt = oRecordSet.Fields.Item("CheckAmt").Value;

        //        //입력금액이 한도금액보다 크면
        //        if (CheckAmt < pAmt)
        //        {
        //            functionReturnValue = false;
        //        }
        //        else
        //        {
        //            functionReturnValue = true;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_CheckAmt_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //    }
        //    finally
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
        //    }
        //    return functionReturnValue;
        //}

        /// <summary>
        ///  지급횟수 계산
        /// </summary>
        /// <param name="pAmt"></param>
        /// <param name="pSchCls"></param>
        /// <returns></returns>
        private int PH_PY301_GetPayCount(string pGovID, string pSchCls, short pDocEntry)
        {
            int functionReturnValue = 0;
            string sQry = null;

            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                sQry = "EXEC PH_PY301_01 '" + pGovID + "','" + pSchCls + "','" + pDocEntry + "'";
                oRecordSet.DoQuery(sQry);
                functionReturnValue = oRecordSet.Fields.Item("PayCount").Value;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_CheckAmt_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return functionReturnValue;
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            int PayCnt = 0;
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "SchCls")
                            {

                                oMat1.FlushToDataSource();

                                //지급횟수 조회
                                //PayCnt = PH_PY301_GetPayCount(oDS_PH_PY301B.GetValue("U_GovID", pVal.Row - 1), oDS_PH_PY301B.GetValue("U_SchCls", pVal.Row - 1), Convert.ToInt16(oDS_PH_PY301A.GetValue("DocEntry", 0)));
                                //oDS_PH_PY301B.SetValue("U_PayCnt", pVal.Row - 1, Convert.ToString(PayCnt));
                                //지급횟수
                                oMat1.LoadFromDataSource();
                            }
                            oMat1.AutoResizeColumns();
                        }
                    }
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

        ///// <summary>
        ///// CLICK 이벤트
        ///// </summary>
        ///// <param name="FormUID">Form UID</param>
        ///// <param name="pVal">ItemEvent 객체</param>
        ///// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        //private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    try
        //    {
        //        if (pVal.Before_Action == true)
        //        {
        //            switch (pVal.ItemUID)
        //            {
        //                case "Mat01":
        //                    if (pVal.Row > 0)
        //                    {
        //                        oLastItemUID = pVal.ItemUID;
        //                        oLastColUID = pVal.ColUID;
        //                        oLastColRow = pVal.Row;

        //                        oMat1.SelectRow(pVal.Row, true, false);
        //                    }
        //                    break;
        //                default:
        //                    oLastItemUID = pVal.ItemUID;
        //                    oLastColUID = "";
        //                    oLastColRow = 0;
        //                    break;
        //            }
        //        }
        //        else if (pVal.Before_Action == false)
        //        {
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //    }
        //}

        ///// <summary>
        ///// VALIDATE 이벤트
        ///// </summary>
        ///// <param name="FormUID">Form UID</param>
        ///// <param name="pVal">ItemEvent 객체</param>
        ///// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        //private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

        //    try
        //    {
        //        if (pVal.Before_Action == true)
        //        {
        //        }
        //        else if (pVal.Before_Action == false)
        //        {
        //            if (pVal.ItemChanged == true)
        //            {
        //                switch (pVal.ItemUID)
        //                {

        //                    case "CntcCode":
        //                        oDS_PH_PY301A.SetValue("U_CntcName", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.VALUE + "'", ""));
        //                        break;

        //                    case "Mat01":

        //                        if (string.IsNullOrEmpty(pVal.ColUID))
        //                        {
        //                            oMat1.FlushToDataSource();
        //                            oMat1.LoadFromDataSource();
        //                            PH_PY301_AddMatrixRow();
        //                        }

        //                        oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //                        oMat1.AutoResizeColumns();
        //                        break;
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //        BubbleEvent = false;
        //    }
        //    finally
        //    {
        //    }
        //}

        ///// <summary>
        ///// MATRIX_LOAD 이벤트
        ///// </summary>
        ///// <param name="FormUID">Form UID</param>
        ///// <param name="pVal">ItemEvent 객체</param>
        ///// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        //private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    try
        //    {
        //        if (pVal.Before_Action == true)
        //        {
        //        }
        //        else if (pVal.Before_Action == false)
        //        {
        //            oMat1.LoadFromDataSource();
        //            PH_PY301_FormItemEnabled();
        //            PH_PY301_AddMatrixRow();
        //            oMat1.AutoResizeColumns();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //    }
        //}

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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY301A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY301B);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
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
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    oMat1.AutoResizeColumns();
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

        ///// <summary>
        ///// CHOOSE_FROM_LIST 이벤트
        ///// </summary>
        ///// <param name="FormUID">Form UID</param>
        ///// <param name="pVal">ItemEvent 객체</param>
        ///// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        //private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

        //    try
        //    {
        //        if (pVal.Before_Action == true)
        //        {
        //        }
        //        else if (pVal.Before_Action == false)
        //        {
        //            //원본 소스(VB6.0 주석처리되어 있음)
        //            //if(pVal.ItemUID == "Code")
        //            //{
        //            //    dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY301A", "Code", "", 0, "", "", "");
        //            //}
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CHOOSE_FROM_LIST_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //    }
        //}

        /// <summary>
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
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
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY301A", "DocEntry"); //접속자 권한에 따른 사업장 보기
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY301_FormItemEnabled();
                            PH_PY301_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        //Case "1293":
                        //  Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
                        case "1281": //문서찾기
                            PH_PY301_FormItemEnabled();
                            PH_PY301_AddMatrixRow();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY301_FormItemEnabled();
                            PH_PY301_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY301_FormItemEnabled();
                            oMat1.AutoResizeColumns();
                            break;
                        case "1293": //행삭제
                            break;
                        case "1287": //복제
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

        ///// <summary>
        ///// FormDataEvent
        ///// </summary>
        ///// <param name="FormUID"></param>
        ///// <param name="BusinessObjectInfo"></param>
        ///// <param name="BubbleEvent"></param>
        //public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //    //string sQry = string.Empty;

        //    SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

        //    try
        //    {
        //        if (BusinessObjectInfo.BeforeAction == true)
        //        {
        //            switch (BusinessObjectInfo.EventType)
        //            {
        //                case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //                    //33
        //                    break;
        //                case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //                    //34
        //                    break;
        //                case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //                    //35
        //                    break;
        //                case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //                    //36
        //                    break;
        //            }
        //        }
        //        else if (BusinessObjectInfo.BeforeAction == false)
        //        {
        //            switch (BusinessObjectInfo.EventType)
        //            {
        //                case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //                    //33
        //                    break;
        //                case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //                    //34
        //                    break;
        //                case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //                    //35
        //                    break;
        //                case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //                    //36
        //                    break;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
        //    }
        //}

        ///// <summary>
        ///// RightClickEvent
        ///// </summary>
        ///// <param name="FormUID"></param>
        ///// <param name="pVal"></param>
        ///// <param name="BubbleEvent"></param>
        //public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        //{
        //    try
        //    {
        //        if (pVal.BeforeAction == true)
        //        {
        //        }
        //        else if (pVal.BeforeAction == false)
        //        {
        //        }

        //        switch (pVal.ItemUID)
        //        {
        //            case "Mat1":
        //                if (pVal.Row > 0)
        //                {
        //                    oLastItemUID = pVal.ItemUID;
        //                    oLastColUID = pVal.ColUID;
        //                    oLastColRow = pVal.Row;
        //                }
        //                break;
        //            default:
        //                oLastItemUID = pVal.ItemUID;
        //                oLastColUID = "";
        //                oLastColRow = 0;
        //                break;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText("Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //    }
        //}
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
//	internal class PH_PY301
//	{
//////********************************************************************************
//////  File           : PH_PY301.cls
//////  Module         : 인사관리 > 기타
//////  Desc           : 학자금신청등록
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		public SAPbouiCOM.Matrix oMat1;

//		private SAPbouiCOM.DBDataSource oDS_PH_PY301A;
//		private SAPbouiCOM.DBDataSource oDS_PH_PY301B;

//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.SP_Screen + "\\PH_PY301.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY301_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY301");
//			PSH_Globals.SBO_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			oForm.DataBrowser.BrowseBy = "DocEntry";

//			oForm.Freeze(true);
//			PH_PY301_CreateItems();
//			PH_PY301_EnableMenus();
//			PH_PY301_SetDocument(oFromDocEntry01);
//			//    Call PH_PY301_FormResize

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
//			PSH_Globals.SBO_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private bool PH_PY301_CreateItems()
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

//			oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oDS_PH_PY301A = oForm.DataSources.DBDataSources("@PH_PY301A");
//			oDS_PH_PY301B = oForm.DataSources.DBDataSources("@PH_PY301B");


//			oMat1 = oForm.Items.Item("Mat01").Specific;

//			oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
//			oMat1.AutoResizeColumns();


//			////----------------------------------------------------------------------------------------------
//			//// 기본사항
//			////----------------------------------------------------------------------------------------------

//			//사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			//    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
//			//    Call SetReDataCombo(oForm, sQry, oCombo)
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;

//			//분기
//			oCombo = oForm.Items.Item("Quarter").Specific;
//			oCombo.ValidValues.Add("", "");
//			oCombo.ValidValues.Add("01", "1/4 혹은 1학기");
//			oCombo.ValidValues.Add("02", "2/4");
//			oCombo.ValidValues.Add("03", "3/4 혹은 2학기");
//			oCombo.ValidValues.Add("04", "4/4");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			oForm.Items.Item("Quarter").DisplayDesc = true;

//			//매트릭스-성별
//			oColumn = oMat1.Columns.Item("Sex");
//			oColumn.ValidValues.Add("", "");
//			oColumn.ValidValues.Add("01", "남자");
//			oColumn.ValidValues.Add("02", "여자");
//			oColumn.DisplayDesc = true;

//			//매트릭스-학교
//			oColumn = oMat1.Columns.Item("SchCls");
//			oColumn.ValidValues.Add("", "");
//			sQry = "            SELECT      T1.U_Code,";
//			sQry = sQry + "                 T1.U_CodeNm";
//			sQry = sQry + "  FROM       [@PS_HR200H] AS T0";
//			sQry = sQry + "                 INNER JOIN";
//			sQry = sQry + "                 [@PS_HR200L] AS T1";
//			sQry = sQry + "                     ON T0.Code = T1.Code";
//			sQry = sQry + "  WHERE      T0.Code = 'P222'";
//			sQry = sQry + "                 AND T1.U_UseYN = 'Y'";
//			sQry = sQry + "  ORDER BY  T1.U_Seq";

//			MDC_SetMod.GP_MatrixSetMatComboList(ref oColumn, ref sQry, ref Convert.ToString(false), ref Convert.ToString(false));

//			//    oColumn.ValidValues.Add "01", "고등학교"
//			//    oColumn.ValidValues.Add "02", "전문대학"
//			//    oColumn.ValidValues.Add "03", "대학교"
//			oColumn.DisplayDesc = true;

//			//매트릭스-학년
//			oColumn = oMat1.Columns.Item("Grade");
//			oColumn.ValidValues.Add("", "");
//			oColumn.ValidValues.Add("01", "1학년");
//			oColumn.ValidValues.Add("02", "2학년");
//			oColumn.ValidValues.Add("03", "3학년");
//			oColumn.ValidValues.Add("04", "4학년");
//			oColumn.ValidValues.Add("05", "5학년");
//			oColumn.DisplayDesc = true;

//			//매트릭스-회차
//			oColumn = oMat1.Columns.Item("Count");
//			oColumn.ValidValues.Add("", "");
//			oColumn.ValidValues.Add("01", "1차");
//			oColumn.ValidValues.Add("02", "2차");
//			oColumn.DisplayDesc = true;



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
//			PH_PY301_CreateItems_Error:

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
//			PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//private void PH_PY301_EnableMenus()
//{

//ERROR: Not supported in C#: OnErrorStatement



//            oForm.EnableMenu("1283", false);
//    // 삭제
//    oForm.EnableMenu("1287", true);
//    // 복제
//    Call oForm.EnableMenu("1286", True)         '// 닫기

//            oForm.EnableMenu("1284", true);
//    // 취소
//    oForm.EnableMenu("1293", true);
//    // 행삭제

//    return;
//PH_PY301_EnableMenus_Error:

//    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//}

//		private void PH_PY301_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY301_FormItemEnabled();
//				PH_PY301_AddMatrixRow();
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY301_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY301_SetDocument_Error:

//			PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//public void PH_PY301_FormItemEnabled()
//{
//    // ERROR: Not supported in C#: OnErrorStatement


//    SAPbouiCOM.ComboBox oCombo = null;

//    oForm.Freeze(true);
//    if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
//    {
//        oForm.Items.Item("CLTCOD").Enabled = true;
//        oForm.Items.Item("CntcCode").Enabled = true;
//        oForm.Items.Item("StdYear").Enabled = true;
//        oForm.Items.Item("Quarter").Enabled = true;
//        oForm.Items.Item("DocEntry").Enabled = false;

//        //폼 DocEntry 세팅
//        PH_PY301_FormClear();

//        //// 접속자에 따른 권한별 사업장 콤보박스세팅
//        MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//        //년도 세팅
//        oDS_PH_PY301A.SetValue("U_StdYear", 0, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYY"));

//        oForm.EnableMenu("1281", true);
//        ////문서찾기
//        oForm.EnableMenu("1282", false);
//        ////문서추가

//    }
//    else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
//    {
//        oForm.Items.Item("CLTCOD").Enabled = true;
//        oForm.Items.Item("CntcCode").Enabled = true;
//        oForm.Items.Item("StdYear").Enabled = true;
//        oForm.Items.Item("Quarter").Enabled = true;
//        oForm.Items.Item("DocEntry").Enabled = true;

//        //// 접속자에 따른 권한별 사업장 콤보박스세팅
//        MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//        oForm.EnableMenu("1281", false);
//        ////문서찾기
//        oForm.EnableMenu("1282", true);
//        ////문서추가

//    }
//    else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
//    {
//        oForm.Items.Item("CLTCOD").Enabled = false;
//        oForm.Items.Item("CntcCode").Enabled = false;
//        oForm.Items.Item("StdYear").Enabled = false;
//        oForm.Items.Item("Quarter").Enabled = false;
//        oForm.Items.Item("DocEntry").Enabled = false;

//        //// 접속자에 따른 권한별 사업장 콤보박스세팅
//        MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);

//        oForm.EnableMenu("1281", true);
//        ////문서찾기
//        oForm.EnableMenu("1282", true);
//        ////문서추가

//    }
//    oForm.Freeze(false);
//    return;
//PH_PY301_FormItemEnabled_Error:

//    oForm.Freeze(false);
//    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//}

//public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//{
//    // ERROR: Not supported in C#: OnErrorStatement


//    string sQry = null;
//    int i = 0;
//    SAPbouiCOM.ComboBox oCombo = null;
//    SAPbobsCOM.Recordset oRecordSet = null;

//    short loopCount = 0;
//    //For Loop 용 (VALIDATE Event에서 사용)
//    string GovID1 = null;
//    //주민등록번호 앞자리(VALIDATE Event에서 사용)
//    string GovID2 = null;
//    //주민등록번호 뒷자리(VALIDATE Event에서 사용)
//    string GovID = null;
//    //주민등록번호 전체(VALIDATE Event에서 사용)
//    string Sex = null;
//    //성별(VALIDATE Event에서 사용)
//    string SchCls = null;
//    //학교(VALIDATE Event에서 사용)
//    short PayCnt = 0;
//    //지급횟수(COMBO_SELECT Event에서 사용)
//    double Tuition = 0;
//    //등록금계(VALIDATE Event에서 사용)
//    double FeeTot = 0;
//    //입학금계(VALIDATE Event에서 사용)
//    double TuiTot = 0;
//    //등록금계(VALIDATE Event에서 사용)
//    double Total = 0;
//    //총계(VALIDATE Event에서 사용)

//    double PreTuition = 0;
//    //등록금 입력 전 데이터

//    oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//    switch (pVal.EventType)
//    {
//        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//            ////1

//            if (pVal.BeforeAction == true)
//            {
//                if (pVal.ItemUID == "1")
//                {
//                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                    {
//                        if (PH_PY301_DataValidCheck() == false)
//                        {
//                            BubbleEvent = false;
//                        }

//                        ////해야할일 작업
//                    }
//                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                    {
//                        if (PH_PY301_DataValidCheck() == false)
//                        {
//                            BubbleEvent = false;
//                        }
//                        ////해야할일 작업

//                    }
//                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//                    {
//                    }
//                }
//            }
//            else if (pVal.BeforeAction == false)
//            {
//                if (pVal.ItemUID == "1")
//                {
//                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                    {
//                        if (pVal.ActionSuccess == true)
//                        {
//                            PH_PY301_FormItemEnabled();
//                            PH_PY301_AddMatrixRow();
//                        }
//                    }
//                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                    {
//                        if (pVal.ActionSuccess == true)
//                        {
//                            PH_PY301_FormItemEnabled();
//                            PH_PY301_AddMatrixRow();
//                        }
//                    }
//                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//                    {
//                        if (pVal.ActionSuccess == true)
//                        {
//                            PH_PY301_FormItemEnabled();
//                        }
//                    }
//                }
//            }
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//            ////2

//            if (pVal.BeforeAction == true)
//            {

//                if (pVal.ItemUID == "Mat01")
//                {

//                    if (pVal.ColUID == "Name" & pVal.CharPressed == Convert.ToDouble("9"))
//                    {

//                        //UPGRADE_WARNING: oMat1.Columns.Item(Name).Cells(pVal.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        if (string.IsNullOrEmpty(oMat1.Columns.Item("Name").Cells.Item(pVal.Row).Specific.VALUE))
//                        {
//                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
//                            BubbleEvent = false;
//                        }

//                    }

//                }
//                else if (pVal.ItemUID == "CntcCode" & pVal.CharPressed == Convert.ToDouble("9"))
//                {

//                    //UPGRADE_WARNING: oForm.Items(CntcCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.VALUE))
//                    {
//                        PSH_Globals.SBO_Application.ActivateMenuItem("7425");
//                        BubbleEvent = false;
//                    }

//                }

//            }
//            else if (pVal.Before_Action == false)
//            {

//            }
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//            ////3
//            switch (pVal.ItemUID)
//            {
//                case "Mat01":
//                    if (pVal.Row > 0)
//                    {
//                        oLastItemUID = pVal.ItemUID;
//                        oLastColUID = pVal.ColUID;
//                        oLastColRow = pVal.Row;
//                    }
//                    break;
//                default:
//                    oLastItemUID = pVal.ItemUID;
//                    oLastColUID = "";
//                    oLastColRow = 0;
//                    break;
//            }
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//            ////4
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//            ////5
//            oForm.Freeze(true);
//            if (pVal.BeforeAction == true)
//            {

//            }
//            else if (pVal.BeforeAction == false)
//            {
//                if (pVal.ItemChanged == true)
//                {
//                    //                    Call PH_PY301_AddMatrixRow

//                    if (pVal.ItemUID == "Mat01")
//                    {
//                        if (pVal.ColUID == "SchCls")
//                        {

//                            oMat1.FlushToDataSource();

//                            //지급횟수 조회
//                            PayCnt = PH_PY301_GetPayCount(oDS_PH_PY301B.GetValue("U_GovID", pVal.Row - 1), oDS_PH_PY301B.GetValue("U_SchCls", pVal.Row - 1), Convert.ToInt16(oDS_PH_PY301A.GetValue("DocEntry", 0)));
//                            oDS_PH_PY301B.SetValue("U_PayCnt", pVal.Row - 1, Convert.ToString(PayCnt));
//                            //지급횟수

//                            oMat1.LoadFromDataSource();

//                        }

//                        oMat1.AutoResizeColumns();
//                    }

//                }
//            }
//            oForm.Freeze(false);
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_CLICK:
//            ////6
//            if (pVal.BeforeAction == true)
//            {
//                switch (pVal.ItemUID)
//                {
//                    case "Mat01":
//                        if (pVal.Row > 0)
//                        {
//                            oMat1.SelectRow(pVal.Row, true, false);
//                        }
//                        break;
//                }

//                switch (pVal.ItemUID)
//                {
//                    case "Mat01":
//                        if (pVal.Row > 0)
//                        {
//                            oLastItemUID = pVal.ItemUID;
//                            oLastColUID = pVal.ColUID;
//                            oLastColRow = pVal.Row;
//                        }
//                        break;
//                    default:
//                        oLastItemUID = pVal.ItemUID;
//                        oLastColUID = "";
//                        oLastColRow = 0;
//                        break;
//                }
//            }
//            else if (pVal.BeforeAction == false)
//            {

//            }
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//            ////7
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//            ////8
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
//            ////9
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//            ////10
//            oForm.Freeze(true);
//            if (pVal.BeforeAction == true)
//            {

//                if (pVal.ItemChanged == true)
//                {

//                }

//            }
//            else if (pVal.BeforeAction == false)
//            {

//                if (pVal.ItemChanged == true)
//                {

//                    switch (pVal.ItemUID)
//                    {

//                        case "CntcCode":
//                            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            //UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oDS_PH_PY301A.SetValue("U_CntcName", 0, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.VALUE + "'"));
//                            break;

//                        case "Mat01":

//                            if (pVal.ColUID == "Name")
//                            {

//                                oMat1.FlushToDataSource();

//                                //UPGRADE_WARNING: oMat1.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                //UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                GovID = MDC_GetData.Get_ReData(ref "U_FamPer", ref "U_FamNam", ref "[@PH_PY001D]", ref "'" + oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE + "'", ref " AND Code = '" + oDS_PH_PY301A.GetValue("U_CntcCode", 0)) + "'");

//                                if (Strings.Right(Strings.Left(GovID, 7), 1) == "1" | Strings.Right(Strings.Left(GovID, 7), 1) == "3" | Strings.Right(Strings.Left(GovID, 7), 1) == "5")
//                                {

//                                    Sex = "01";

//                                }
//                                else if (Strings.Right(Strings.Left(GovID, 7), 1) == "2" | Strings.Right(Strings.Left(GovID, 7), 1) == "4" | Strings.Right(Strings.Left(GovID, 7), 1) == "6")
//                                {

//                                    Sex = "02";

//                                }

//                                GovID1 = Strings.Left(GovID, 6);
//                                GovID2 = Strings.Right(GovID, 7);
//                                GovID = GovID1 + "-" + GovID2;

//                                //                                PayCnt = MDC_GetData.Get_ReData("COUNT(*)", "U_GovID", "[@PH_PY301B]", "'" & GovID & "'") '주민등록번호를 "-"을 포함하여 저장하기 때문에 지급횟수 조회 로직은 여기에 있어야 함

//                                //UPGRADE_WARNING: oMat1.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                oDS_PH_PY301B.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE);
//                                oDS_PH_PY301B.SetValue("U_GovID", pVal.Row - 1, GovID);
//                                //주민등록번호
//                                oDS_PH_PY301B.SetValue("U_Sex", pVal.Row - 1, Sex);
//                                //성별
//                                //                                Call oDS_PH_PY301B.setValue("U_PayCnt", pVal.Row - 1, PayCnt) '지급횟수
//                                oDS_PH_PY301B.SetValue("U_PayYN", pVal.Row - 1, "N");
//                                //지급완료여부

//                                oMat1.LoadFromDataSource();

//                                PH_PY301_AddMatrixRow();

//                                //입학금 입력 시
//                            }
//                            else if (pVal.ColUID == "EntFee")
//                            {

//                                oMat1.FlushToDataSource();

//                                //학교선택을 하지 않으면 에러 메시지 출력
//                                if (string.IsNullOrEmpty(oDS_PH_PY301B.GetValue("U_SchCls", pVal.Row - 1)))
//                                {

//                                    MDC_Com.MDC_GF_Message(ref "학교를 먼저 선택하십시오.", ref "E");
//                                    oDS_PH_PY301B.SetValue("U_EntFee", pVal.Row - 1, Convert.ToString(0));
//                                    BubbleEvent = false;

//                                }

//                                //입학금 합계 계산
//                                for (loopCount = 0; loopCount <= oMat1.RowCount - 1; loopCount++)
//                                {

//                                    FeeTot = FeeTot + Convert.ToDouble(oDS_PH_PY301B.GetValue("U_EntFee", loopCount));

//                                }
//                                oMat1.LoadFromDataSource();

//                                oDS_PH_PY301A.SetValue("U_FeeTot", 0, Convert.ToString(FeeTot));

//                                TuiTot = Convert.ToDouble(oDS_PH_PY301A.GetValue("U_TuiTot", 0));
//                                Total = FeeTot + TuiTot;

//                                oDS_PH_PY301A.SetValue("U_Total", 0, Convert.ToString(Total));


//                                //등록금 입력 시
//                            }
//                            else if (pVal.ColUID == "Tuition")
//                            {

//                                PreTuition = Convert.ToDouble(oDS_PH_PY301B.GetValue("U_Tuition", pVal.Row - 1));

//                                oMat1.FlushToDataSource();

//                                //학교선택을 하지 않으면 에러 메시지 출력
//                                if (string.IsNullOrEmpty(oDS_PH_PY301B.GetValue("U_SchCls", pVal.Row - 1)))
//                                {

//                                    MDC_Com.MDC_GF_Message(ref "학교를 먼저 선택하십시오.", ref "E");
//                                    oDS_PH_PY301B.SetValue("U_Tuition", pVal.Row - 1, Convert.ToString(0));
//                                    BubbleEvent = false;

//                                }

//                                //한도금액 체크
//                                SchCls = oDS_PH_PY301B.GetValue("U_SchCls", pVal.Row - 1));
//                                Tuition = Convert.ToDouble(oDS_PH_PY301B.GetValue("U_Tuition", pVal.Row - 1)));

//                                //고등학교 이외만 체크
//                                if (SchCls > "02")
//                                {
//                                    if (PH_PY301_CheckAmt(Tuition, SchCls) == false)
//                                    {

//                                        MDC_Com.MDC_GF_Message(ref "등록금이 한도금액을 초과하였습니다. 확인하십시오.", ref "E");
//                                        oDS_PH_PY301B.SetValue("U_Tuition", pVal.Row - 1, Convert.ToString(PreTuition));
//                                        //이전 데이터로 회귀
//                                        BubbleEvent = false;

//                                    }
//                                }

//                                //등록금 합계 계산
//                                for (loopCount = 0; loopCount <= oMat1.RowCount - 1; loopCount++)
//                                {

//                                    TuiTot = TuiTot + Convert.ToDouble(oDS_PH_PY301B.GetValue("U_Tuition", loopCount));

//                                }
//                                oMat1.LoadFromDataSource();

//                                oDS_PH_PY301A.SetValue("U_TuiTot", 0, Convert.ToString(TuiTot));

//                                FeeTot = Convert.ToDouble(oDS_PH_PY301A.GetValue("U_FeeTot", 0));
//                                Total = FeeTot + TuiTot;

//                                oDS_PH_PY301A.SetValue("U_Total", 0, Convert.ToString(Total));

//                            }

//                            oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                            oMat1.AutoResizeColumns();
//                            break;

//                    }

//                }

//            }
//            oForm.Freeze(false);
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//            ////11
//            if (pVal.BeforeAction == true)
//            {
//            }
//            else if (pVal.BeforeAction == false)
//            {
//                oMat1.LoadFromDataSource();

//                PH_PY301_FormItemEnabled();
//                PH_PY301_AddMatrixRow();
//                oMat1.AutoResizeColumns();

//            }
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
//            ////12
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
//            ////16
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//            ////17
//            if (pVal.BeforeAction == true)
//            {
//            }
//            else if (pVal.BeforeAction == false)
//            {
//                SubMain.RemoveForms(oFormUniqueID);
//                //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                oForm = null;
//                //UPGRADE_NOTE: oDS_PH_PY301A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                oDS_PH_PY301A = null;
//                //UPGRADE_NOTE: oDS_PH_PY301B 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                oDS_PH_PY301B = null;

//                //UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                oMat1 = null;

//            }
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//            ////18
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//            ////19
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
//            ////20
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//            ////21
//            if (pVal.BeforeAction == true)
//            {

//            }
//            else if (pVal.BeforeAction == false)
//            {

//                oMat1.AutoResizeColumns();

//            }
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
//            ////22
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
//            ////23
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//            ////27
//            if (pVal.BeforeAction == true)
//            {

//            }
//            else if (pVal.Before_Action == false)
//            {
//                //                If pVal.ItemUID = "Code" Then
//                //                    Call MDC_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY301A", "Code")
//                //                End If
//            }
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
//            ////37
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
//            ////38
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_Drag:
//            ////39
//            break;

//    }

//    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oCombo = null;
//    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oRecordSet = null;

//    return;
//Raise_FormItemEvent_Error:
//    ///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//    oForm.Freeze((false));
//    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oCombo = null;
//    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oRecordSet = null;
//    PSH_Globals.SBO_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//}


//public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
//{
//    int i = 0;
//    // ERROR: Not supported in C#: OnErrorStatement


//    short loopCount = 0;
//    double FeeTot = 0;
//    double TuiTot = 0;
//    double Total = 0;

//    oForm.Freeze(true);

//    if ((pVal.BeforeAction == true))
//    {
//        switch (pVal.MenuUID)
//        {
//            case "1283":
//                if (PSH_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
//                {
//                    BubbleEvent = false;
//                    return;
//                }
//                break;
//            case "1284":
//                break;
//            case "1286":
//                break;
//            case "1293":
//                break;
//            case "1281":
//                break;
//            case "1282":
//                break;
//            case "1288":
//            case "1289":
//            case "1290":
//            case "1291":
//                MDC_SetMod.AuthorityCheck(ref oForm, ref "CLTCOD", ref "@PH_PY301A", ref "DocEntry");
//                ////접속자 권한에 따른 사업장 보기
//                break;
//        }
//    }
//    else if ((pVal.BeforeAction == false))
//    {
//        switch (pVal.MenuUID)
//        {
//            case "1283":
//                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                PH_PY301_FormItemEnabled();
//                PH_PY301_AddMatrixRow();
//                break;
//            case "1284":
//                break;
//            case "1286":
//                break;
//            //            Case "1293":
//            //                Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
//            case "1281":
//                ////문서찾기
//                PH_PY301_FormItemEnabled();
//                PH_PY301_AddMatrixRow();
//                oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                break;
//            case "1282":
//                ////문서추가
//                PH_PY301_FormItemEnabled();
//                PH_PY301_AddMatrixRow();
//                break;
//            case "1288":
//            case "1289":
//            case "1290":
//            case "1291":
//                PH_PY301_FormItemEnabled();
//                break;
//            case "1293":
//                //// 행삭제

//                if (oMat1.RowCount != oMat1.VisualRowCount)
//                {
//                    oMat1.FlushToDataSource();

//                    while ((i <= oDS_PH_PY301B.Size - 1))
//                    {
//                        if (string.IsNullOrEmpty(oDS_PH_PY301B.GetValue("U_LineNum", i)))
//                        {
//                            oDS_PH_PY301B.RemoveRecord((i));
//                            i = 0;
//                        }
//                        else
//                        {
//                            i = i + 1;
//                        }
//                    }

//                    for (i = 0; i <= oDS_PH_PY301B.Size; i++)
//                    {
//                        oDS_PH_PY301B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//                    }

//                    oMat1.LoadFromDataSource();
//                }
//                PH_PY301_AddMatrixRow();

//                //합계 재 계산
//                oMat1.FlushToDataSource();
//                for (loopCount = 0; loopCount <= oMat1.RowCount - 1; loopCount++)
//                {

//                    FeeTot = FeeTot + Convert.ToDouble(oDS_PH_PY301B.GetValue("U_EntFee", loopCount));
//                    TuiTot = TuiTot + Convert.ToDouble(oDS_PH_PY301B.GetValue("U_Tuition", loopCount));

//                }

//                oMat1.LoadFromDataSource();

//                Total = FeeTot + TuiTot;

//                oDS_PH_PY301A.SetValue("U_FeeTot", 0, Convert.ToString(FeeTot));
//                oDS_PH_PY301A.SetValue("U_TuiTot", 0, Convert.ToString(TuiTot));
//                oDS_PH_PY301A.SetValue("U_Total", 0, Convert.ToString(Total));
//                break;
//            //합계 재 계산

//            //복제
//            case "1287":

//                oForm.Freeze(true);
//                oDS_PH_PY301A.SetValue("DocEntry", 0, "");

//                for (i = 0; i <= oMat1.VisualRowCount - 1; i++)
//                {
//                    oMat1.FlushToDataSource();
//                    oDS_PH_PY301B.SetValue("DocEntry", i, "");
//                    oDS_PH_PY301B.SetValue("U_PayYN", i, "N");
//                    oMat1.LoadFromDataSource();
//                }


//                oForm.Items.Item("Quarter").Enabled = true;
//                oForm.Freeze(false);
//                break;

//        }
//    }
//    oForm.Freeze(false);
//    return;
//Raise_FormMenuEvent_Error:
//    oForm.Freeze(false);
//    PSH_Globals.SBO_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//}

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


//			PSH_Globals.SBO_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//		}

//		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pVal.BeforeAction == true) {
//			} else if (pVal.BeforeAction == false) {
//			}
//			switch (pVal.ItemUID) {
//				case "Mat01":
//					if (pVal.Row > 0) {
//						oLastItemUID = pVal.ItemUID;
//						oLastColUID = pVal.ColUID;
//						oLastColRow = pVal.Row;
//					}
//					break;
//				default:
//					oLastItemUID = pVal.ItemUID;
//					oLastColUID = "";
//					oLastColRow = 0;
//					break;
//			}
//			return;
//			Raise_RightClickEvent_Error:

//			PSH_Globals.SBO_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//public void PH_PY301_AddMatrixRow()
//{
//    int oRow = 0;

//    // ERROR: Not supported in C#: OnErrorStatement


//    oForm.Freeze(true);

//    ////[Mat1]
//    oMat1.FlushToDataSource();
//    oRow = oMat1.VisualRowCount;

//    if (oMat1.VisualRowCount > 0)
//    {
//        if (!string.IsNullOrEmpty(oDS_PH_PY301B.GetValue("U_Name", oRow - 1))))
//        {
//            if (oDS_PH_PY301B.Size <= oMat1.VisualRowCount)
//            {
//                oDS_PH_PY301B.InsertRecord((oRow));
//            }
//            oDS_PH_PY301B.Offset = oRow;
//            oDS_PH_PY301B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//            oDS_PH_PY301B.SetValue("U_Name", oRow, "");
//            oDS_PH_PY301B.SetValue("U_GovID", oRow, "");
//            oDS_PH_PY301B.SetValue("U_Sex", oRow, "");
//            oDS_PH_PY301B.SetValue("U_SchCls", oRow, "");
//            oDS_PH_PY301B.SetValue("U_SchName", oRow, "");
//            oDS_PH_PY301B.SetValue("U_Grade", oRow, "");
//            oDS_PH_PY301B.SetValue("U_EntFee", oRow, Convert.ToString(0));
//            oDS_PH_PY301B.SetValue("U_Tuition", oRow, Convert.ToString(0));
//            oDS_PH_PY301B.SetValue("U_Count", oRow, "");
//            oDS_PH_PY301B.SetValue("U_PayCnt", oRow, "");
//            oDS_PH_PY301B.SetValue("U_PayYN", oRow, "");
//            oMat1.LoadFromDataSource();
//        }
//        else
//        {
//            oDS_PH_PY301B.Offset = oRow - 1;
//            oDS_PH_PY301B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//            oDS_PH_PY301B.SetValue("U_Name", oRow - 1, "");
//            oDS_PH_PY301B.SetValue("U_GovID", oRow - 1, "");
//            oDS_PH_PY301B.SetValue("U_Sex", oRow - 1, "");
//            oDS_PH_PY301B.SetValue("U_SchCls", oRow - 1, "");
//            oDS_PH_PY301B.SetValue("U_SchName", oRow - 1, "");
//            oDS_PH_PY301B.SetValue("U_Grade", oRow - 1, "");
//            oDS_PH_PY301B.SetValue("U_EntFee", oRow - 1, Convert.ToString(0));
//            oDS_PH_PY301B.SetValue("U_Tuition", oRow - 1, Convert.ToString(0));
//            oDS_PH_PY301B.SetValue("U_Count", oRow - 1, "");
//            oDS_PH_PY301B.SetValue("U_PayCnt", oRow, "");
//            oDS_PH_PY301B.SetValue("U_PayYN", oRow - 1, "");
//            oMat1.LoadFromDataSource();
//        }
//    }
//    else if (oMat1.VisualRowCount == 0)
//    {
//        oDS_PH_PY301B.Offset = oRow;
//        oDS_PH_PY301B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//        oDS_PH_PY301B.SetValue("U_Name", oRow, "");
//        oDS_PH_PY301B.SetValue("U_GovID", oRow, "");
//        oDS_PH_PY301B.SetValue("U_Sex", oRow, "");
//        oDS_PH_PY301B.SetValue("U_SchCls", oRow, "");
//        oDS_PH_PY301B.SetValue("U_SchName", oRow, "");
//        oDS_PH_PY301B.SetValue("U_Grade", oRow, "");
//        oDS_PH_PY301B.SetValue("U_EntFee", oRow, Convert.ToString(0));
//        oDS_PH_PY301B.SetValue("U_Tuition", oRow, Convert.ToString(0));
//        oDS_PH_PY301B.SetValue("U_Count", oRow, "");
//        oDS_PH_PY301B.SetValue("U_PayCnt", oRow, "");
//        oDS_PH_PY301B.SetValue("U_PayYN", oRow, "");
//        oMat1.LoadFromDataSource();
//    }

//    oForm.Freeze(false);
//    return;
//PH_PY301_AddMatrixRow_Error:
//    oForm.Freeze(false);
//    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//}

//public void PH_PY301_FormClear()
//{
//    // ERROR: Not supported in C#: OnErrorStatement

//    string DocEntry = null;
//    //UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//    DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY301'", ref "");
//    if (Convert.ToDouble(DocEntry) == 0)
//    {
//        //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//    }
//    else
//    {
//        //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//    }
//    return;
//PH_PY301_FormClear_Error:
//    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//}

//public bool PH_PY301_DataValidCheck()
//{
//    bool functionReturnValue = false;
//    // ERROR: Not supported in C#: OnErrorStatement

//    functionReturnValue = false;
//    int i = 0;
//    string sQry = null;
//    SAPbobsCOM.Recordset oRecordSet = null;

//    string CLTCOD = null;
//    string StdYear = null;
//    string Quarter = null;
//    string Count = null;

//    oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);




//    //사업장
//    if (string.IsNullOrEmpty(oDS_PH_PY301A.GetValue("U_CLTCOD", 0))))
//    {
//        PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        functionReturnValue = false;
//        return functionReturnValue;
//    }

//    //년도
//    if (string.IsNullOrEmpty(oDS_PH_PY301A.GetValue("U_StdYear", 0))))
//    {
//        PSH_Globals.SBO_Application.SetStatusBarMessage("년도는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        oForm.Items.Item("StdYear").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        functionReturnValue = false;
//        return functionReturnValue;
//    }

//    //사번
//    if (string.IsNullOrEmpty(oDS_PH_PY301A.GetValue("U_CntcCode", 0))))
//    {
//        PSH_Globals.SBO_Application.SetStatusBarMessage("사번은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        functionReturnValue = false;
//        return functionReturnValue;
//    }

//    //분기
//    if (string.IsNullOrEmpty(oDS_PH_PY301A.GetValue("U_Quarter", 0))))
//    {
//        PSH_Globals.SBO_Application.SetStatusBarMessage("분기는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        oForm.Items.Item("Quarter").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        functionReturnValue = false;
//        return functionReturnValue;
//    }

//    CLTCOD = oDS_PH_PY301A.GetValue("U_CLTCOD", 0));
//    StdYear = oDS_PH_PY301A.GetValue("U_StdYear", 0));
//    Quarter = oDS_PH_PY301A.GetValue("U_Quarter", 0));

//    //라인
//    if (oMat1.VisualRowCount > 1)
//    {
//        for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
//        {

//            //학교
//            //UPGRADE_WARNING: oMat1.Columns(SchCls).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            if (string.IsNullOrEmpty(oMat1.Columns.Item("SchCls").Cells.Item(i).Specific.VALUE))
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("학교는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                oMat1.Columns.Item("SchCls").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                functionReturnValue = false;
//                return functionReturnValue;
//            }

//            //학교명
//            //UPGRADE_WARNING: oMat1.Columns(SchName).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            if (string.IsNullOrEmpty(oMat1.Columns.Item("SchName").Cells.Item(i).Specific.VALUE))
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("학교명은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                oMat1.Columns.Item("SchName").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                functionReturnValue = false;
//                return functionReturnValue;
//            }

//            //학년
//            //UPGRADE_WARNING: oMat1.Columns(Grade).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            if (string.IsNullOrEmpty(oMat1.Columns.Item("Grade").Cells.Item(i).Specific.VALUE))
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("학년은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                oMat1.Columns.Item("Grade").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                functionReturnValue = false;
//                return functionReturnValue;
//            }

//            //회차
//            //UPGRADE_WARNING: oMat1.Columns(Count).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            if (string.IsNullOrEmpty(oMat1.Columns.Item("Count").Cells.Item(i).Specific.VALUE))
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("회차는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                oMat1.Columns.Item("Count").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                functionReturnValue = false;
//                return functionReturnValue;
//            }

//            //UPGRADE_WARNING: oMat1.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Count = oMat1.Columns.Item("Count").Cells.Item(i).Specific.VALUE;

//            sQry = "Select Cnt = Count(*) From [@PH_PY301A] a Inner Join [@PH_PY301B] b On a.DocEntry = b.DocEntry and a.Canceled = 'N' ";
//            sQry = sQry + " Where a.U_CLTCOD = '" + CLTCOD + "' And a.U_StdYear = '" + StdYear + "' and a.U_Quarter = '" + Quarter + "' ";
//            sQry = sQry + " And b.U_Count = '" + Count + "' and b.U_PayYN = 'Y'";

//            oRecordSet.DoQuery(sQry);

//            if (oRecordSet.Fields.Item(0).Value > 0)
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("지급완료처리가 되어 추가/수정을 할 수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                functionReturnValue = false;
//                return functionReturnValue;
//            }

//        }
//    }
//    else
//    {
//        PSH_Globals.SBO_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        functionReturnValue = false;
//        return functionReturnValue;
//    }





//    oMat1.FlushToDataSource();
//    //// Matrix 마지막 행 삭제(DB 저장시)
//    if (oDS_PH_PY301B.Size > 1)
//        oDS_PH_PY301B.RemoveRecord((oDS_PH_PY301B.Size - 1));

//    oMat1.LoadFromDataSource();

//    functionReturnValue = true;
//    return functionReturnValue;


//    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oRecordSet = null;
//PH_PY301_DataValidCheck_Error:


//    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oRecordSet = null;
//    functionReturnValue = false;
//    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//    return functionReturnValue;
//}

//		private void PH_PY301_MTX01()
//		{

//			////메트릭스에 데이터 로드

//			int i = 0;
//			string sQry = null;

//			string Param01 = null;
//			string Param02 = null;
//			string Param03 = null;
//			string Param04 = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param01 = oForm.Items.Item("Param01").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param02 = oForm.Items.Item("Param01").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param03 = oForm.Items.Item("Param01").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param04 = oForm.Items.Item("Param01").Specific.VALUE;

//			sQry = "SELECT 10";
//			oRecordSet.DoQuery(sQry);

//			oMat1.Clear();
//			oMat1.FlushToDataSource();
//			oMat1.LoadFromDataSource();

//			if ((oRecordSet.RecordCount == 0)) {
//				MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
//				goto PH_PY301_MTX01_Exit;
//			}

//			SAPbouiCOM.ProgressBar ProgressBar01 = null;
//			ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

//			for (i = 0; i <= oRecordSet.RecordCount - 1; i++) {
//				if (i != 0) {
//					oDS_PH_PY301B.InsertRecord((i));
//				}
//				oDS_PH_PY301B.Offset = i;
//				oDS_PH_PY301B.SetValue("U_COL01", i, oRecordSet.Fields.Item(0).Value);
//				oDS_PH_PY301B.SetValue("U_COL02", i, oRecordSet.Fields.Item(1).Value);
//				oRecordSet.MoveNext();
//				ProgressBar01.Value = ProgressBar01.Value + 1;
//				ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
//			}
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
//			PH_PY301_MTX01_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			if ((ProgressBar01 != null)) {
//				ProgressBar01.Stop();
//			}
//			return;
//			PH_PY301_MTX01_Error:
//			ProgressBar01.Stop();
//			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgressBar01 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY301_Validate(string ValidateType)
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = true;
//			object i = null;
//			int j = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY301A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY301A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
//				PSH_Globals.SBO_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY301_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY301_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY301_Validate_Error:
//			functionReturnValue = false;
//			PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}



//private short PH_PY301_GetPayCount(string pGovID, string pSchCls, short pDocEntry)
//{
//    short functionReturnValue = 0;
//    //******************************************************************************
//    //Function ID : PH_PY301_GetPayCount()
//    //해당모듈 : PH_PY301
//    //기능 : 지급횟수 계산
//    //인수 : pGovID:주민등록번호, pSchCls:학교구분(고등학교:01, 전문대학:02, 대학교:03), pDocEntry:문서번호
//    //반환값 : 지급횟수
//    //특이사항 : 없음
//    //******************************************************************************
//    // ERROR: Not supported in C#: OnErrorStatement


//    short loopCount = 0;
//    string sQry = null;
//    object CheckAmt = null;

//    SAPbobsCOM.Recordset oRecordSet = null;
//    oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//    sQry = "EXEC PH_PY301_01 '" + pGovID + "','" + pSchCls + "','" + pDocEntry + "'";

//    oRecordSet.DoQuery(sQry);

//    //UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//    functionReturnValue = oRecordSet.Fields.Item("PayCount").Value;
//    return functionReturnValue;
//PH_PY301_GetPayCount_Error:

//    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oRecordSet = null;
//    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_GetPayCount_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//    return functionReturnValue;
//}
//	}
//}
