using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 승호등록작업
    /// </summary>
    internal class PH_PY006 : PSH_BaseClass
    {
        public string oFormUniqueID;
        //public SAPbouiCOM.Form oForm;

        public SAPbouiCOM.Matrix oMat1;

        private SAPbouiCOM.DBDataSource oDS_PH_PY006A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY006B;

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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY006.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY006_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY006");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                oForm.Visible = true;
                PH_PY006_CreateItems();
                PH_PY006_EnableMenus();
                PH_PY006_SetDocument(oFromDocEntry01);
                //    Call PH_PY006_FormResize
            }
            catch(Exception ex)
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
        private void PH_PY006_CreateItems()
        {
            string sQry = null;
            //int i = 0;

            //SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.ComboBox oCombo = null;
            SAPbouiCOM.Column oColumn = null;
            //SAPbouiCOM.Columns oColumns = null;
            SAPbobsCOM.Recordset oRecordSet = null;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            oForm.Freeze(true);

            try
            {
                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oDS_PH_PY006A = oForm.DataSources.DBDataSources.Item("@PH_PY006A");
                oDS_PH_PY006B = oForm.DataSources.DBDataSources.Item("@PH_PY006B");

                oMat1 = oForm.Items.Item("Mat01").Specific;

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMat1.AutoResizeColumns();

                //사업장
                oCombo = oForm.Items.Item("CLTCOD").Specific;
                //    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
                //    Call SetReDataCombo(oForm, sQry, oCombo)
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //직원구분
                oCombo = oForm.Items.Item("JIGTYP").Specific;
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P126' AND U_UseYN= 'Y' ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "");
                oForm.Items.Item("JIGTYP").DisplayDesc = true;

                //승호구분
                oCombo = oForm.Items.Item("Gubun").Specific;
                oCombo.ValidValues.Add("10", "금액승호(전문직)");
                oCombo.ValidValues.Add("20", "호봉승호(사무기술직)");
                oForm.Items.Item("JIGTYP").DisplayDesc = true;


                //매트릭스-부서
                oColumn = oMat1.Columns.Item("TeamCode");
                sQry = "        SELECT      T1.U_Code,";
                sQry = sQry + "             T1.U_CodeNm";
                sQry = sQry + " FROM        [@PS_HR200H] AS T0";
                sQry = sQry + "             INNER JOIN";
                sQry = sQry + "             [@PS_HR200L] AS T1";
                sQry = sQry + "                 ON T0.Code = T1.Code";
                sQry = sQry + " WHERE       T0.Code = '1'";
                sQry = sQry + "             AND T1.U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    T1.U_Seq";

                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }

                //    Call MDC_SetMod.GP_MatrixSetMatComboList(oColumn, sQry, False, False)
                oColumn.DisplayDesc = true;

                //매트릭스-담당
                oColumn = oMat1.Columns.Item("RspCode");
                sQry = "        SELECT      T1.U_Code,";
                sQry = sQry + "             T1.U_CodeNm";
                sQry = sQry + " FROM        [@PS_HR200H] AS T0";
                sQry = sQry + "             INNER JOIN";
                sQry = sQry + "             [@PS_HR200L] AS T1";
                sQry = sQry + "                 ON T0.Code = T1.Code";
                sQry = sQry + " WHERE       T0.Code = '1'";
                sQry = sQry + "             AND T1.U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    T1.U_Seq";

                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }

                //    Call MDC_SetMod.GP_MatrixSetMatComboList(oColumn, sQry, False, False)
                oColumn.DisplayDesc = true;

                //매트릭스-반
                oColumn = oMat1.Columns.Item("ClsCode");
                sQry = "        SELECT      T1.U_Code,";
                sQry = sQry + "             T1.U_CodeNm";
                sQry = sQry + " FROM        [@PS_HR200H] AS T0";
                sQry = sQry + "             INNER JOIN";
                sQry = sQry + "             [@PS_HR200L] AS T1";
                sQry = sQry + "                 ON T0.Code = T1.Code";
                sQry = sQry + " WHERE       T0.Code = '9'";
                sQry = sQry + "             AND T1.U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    T1.U_Seq";

                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }

                //    Call MDC_SetMod.GP_MatrixSetMatComboList(oColumn, sQry, False, False)
                oColumn.DisplayDesc = true;

                //매트릭스-직급
                oColumn = oMat1.Columns.Item("JIGCOD");
                sQry = "        SELECT      T1.U_Code,";
                sQry = sQry + "             T1.U_CodeNm";
                sQry = sQry + " FROM        [@PS_HR200H] AS T0";
                sQry = sQry + "             INNER JOIN";
                sQry = sQry + "             [@PS_HR200L] AS T1";
                sQry = sQry + "                 ON T0.Code = T1.Code";
                sQry = sQry + "  WHERE      T0.Code = 'P129'";
                sQry = sQry + "             AND T1.U_UseYN = 'Y'";
                sQry = sQry + "  ORDER BY   T1.U_Seq";

                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }

                //    Call MDC_SetMod.GP_MatrixSetMatComboList(oColumn, sQry, False, False)
                oColumn.DisplayDesc = true;

                //승호대상자 여부
                oColumn = oMat1.Columns.Item("APPLYYN");
                oColumn.ValidValues.Add("Y", "승호대상자");
                oColumn.ValidValues.Add("N", "승호제외자");
                oColumn.DisplayDesc = true;

                //승호대상자 여부
                oColumn = oMat1.Columns.Item("PeakYN");
                oColumn.ValidValues.Add("Y", "대상");
                oColumn.ValidValues.Add("N", "비대상");
                oColumn.DisplayDesc = true;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY006_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);

                //메모리 해제
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCombo);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY006_EnableMenus()
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
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY006_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY006_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PH_PY006_FormItemEnabled();
                    PH_PY006_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY006_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY006_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY006_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            //SAPbouiCOM.ComboBox oCombo = null;

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("UPAMT").Enabled = true;
                    oForm.Items.Item("appNum").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;

                    //폼 DocEntry 세팅
                    PH_PY006_FormClear();

                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    //년월 세팅
                    oDS_PH_PY006A.SetValue("U_YM", 0, DateTime.Now.ToString("yyyyMM"));
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("UPAMT").Enabled = true;
                    oForm.Items.Item("JIGTYP").Enabled = true;
                    oForm.Items.Item("appNum").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;

                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("YM").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oForm.Items.Item("UPAMT").Enabled = false;
                    oForm.Items.Item("JIGTYP").Enabled = false;
                    oForm.Items.Item("appNum").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;

                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);

                    //oForm.EnableMenu("1281", True)     '//문서찾기
                    //oForm.EnableMenu("1282", True)     '//문서추가
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY006_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Matirx 행 추가
        /// </summary>
        private void PH_PY006_AddMatrixRow()
        {
            int oRow = 0;

            try
            {
                oForm.Freeze(true);
                //[Mat1]
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                //    If oMat1.VisualRowCount > 0 Then
                //        If Trim(oDS_PH_PY006B.GetValue("U_Name", oRow - 1)) <> "" Then
                //            If oDS_PH_PY006B.Size <= oMat1.VisualRowCount Then
                //                oDS_PH_PY006B.InsertRecord (oRow)
                //            End If
                //            oDS_PH_PY006B.Offset = oRow
                //            oDS_PH_PY006B.setValue "U_LineNum", oRow, oRow + 1
                //            oDS_PH_PY006B.setValue "U_Name", oRow, ""
                //            oDS_PH_PY006B.setValue "U_GovID", oRow, ""
                //            oDS_PH_PY006B.setValue "U_Sex", oRow, ""
                //            oDS_PH_PY006B.setValue "U_SchCls", oRow, ""
                //            oDS_PH_PY006B.setValue "U_SchName", oRow, ""
                //            oDS_PH_PY006B.setValue "U_Grade", oRow, ""
                //            oDS_PH_PY006B.setValue "U_EntFee", oRow, 0
                //            oDS_PH_PY006B.setValue "U_Tuition", oRow, 0
                //            oDS_PH_PY006B.setValue "U_Count", oRow, ""
                //            oDS_PH_PY006B.setValue "U_PayCnt", oRow, ""
                //            oDS_PH_PY006B.setValue "U_PayYN", oRow, ""
                //            oMat1.LoadFromDataSource
                //        Else
                //            oDS_PH_PY006B.Offset = oRow - 1
                //            oDS_PH_PY006B.setValue "U_LineNum", oRow - 1, oRow
                //            oDS_PH_PY006B.setValue "U_Name", oRow - 1, ""
                //            oDS_PH_PY006B.setValue "U_GovID", oRow - 1, ""
                //            oDS_PH_PY006B.setValue "U_Sex", oRow - 1, ""
                //            oDS_PH_PY006B.setValue "U_SchCls", oRow - 1, ""
                //            oDS_PH_PY006B.setValue "U_SchName", oRow - 1, ""
                //            oDS_PH_PY006B.setValue "U_Grade", oRow - 1, ""
                //            oDS_PH_PY006B.setValue "U_EntFee", oRow - 1, 0
                //            oDS_PH_PY006B.setValue "U_Tuition", oRow - 1, 0
                //            oDS_PH_PY006B.setValue "U_Count", oRow - 1, ""
                //            oDS_PH_PY006B.setValue "U_PayCnt", oRow, ""
                //            oDS_PH_PY006B.setValue "U_PayYN", oRow - 1, ""
                //            oMat1.LoadFromDataSource
                //        End If
                //    ElseIf oMat1.VisualRowCount = 0 Then
                //        oDS_PH_PY006B.Offset = oRow
                //        oDS_PH_PY006B.setValue "U_LineNum", oRow, oRow + 1
                //        oDS_PH_PY006B.setValue "U_Name", oRow, ""
                //        oDS_PH_PY006B.setValue "U_GovID", oRow, ""
                //        oDS_PH_PY006B.setValue "U_Sex", oRow, ""
                //        oDS_PH_PY006B.setValue "U_SchCls", oRow, ""
                //        oDS_PH_PY006B.setValue "U_SchName", oRow, ""
                //        oDS_PH_PY006B.setValue "U_Grade", oRow, ""
                //        oDS_PH_PY006B.setValue "U_EntFee", oRow, 0
                //        oDS_PH_PY006B.setValue "U_Tuition", oRow, 0
                //        oDS_PH_PY006B.setValue "U_Count", oRow, ""
                //        oDS_PH_PY006B.setValue "U_PayCnt", oRow, ""
                //        oDS_PH_PY006B.setValue "U_PayYN", oRow, ""
                //        oMat1.LoadFromDataSource
                //    End If
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY006_AddMatrixRow_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY006_FormClear()
        {
            string DocEntry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY006'", "");
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY006_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY006_DataValidCheck()
        {
            bool functionReturnValue = false;
            
            functionReturnValue = false;
            //int i = 0;
            //string sQry = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //사업장
                if (string.IsNullOrEmpty(oDS_PH_PY006A.GetValue("U_CLTCOD", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                //년월
                if (string.IsNullOrEmpty(oDS_PH_PY006A.GetValue("U_YM", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("년월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                //승호일
                if (string.IsNullOrEmpty(oDS_PH_PY006A.GetValue("U_DocDate", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("승호일은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                //승호금액
                if (string.IsNullOrEmpty(oDS_PH_PY006A.GetValue("U_UPAMT", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("승호금액은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("UPAMT").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                //직원구분
                if (string.IsNullOrEmpty(oDS_PH_PY006A.GetValue("U_JIGTYP", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("직원구분은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("JIGTYP").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                //발령번호
                if (string.IsNullOrEmpty(oDS_PH_PY006A.GetValue("U_appNum", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("발령번호는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("appNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                //라인
                if (oMat1.VisualRowCount >= 1)
                {
                    //        For i = 1 To oMat1.VisualRowCount - 1
                    //
                    //        Next
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                oMat1.FlushToDataSource();
                //// Matrix 마지막 행 삭제(DB 저장시)
                //    If oDS_PH_PY006B.Size > 1 Then oDS_PH_PY006B.RemoveRecord (oDS_PH_PY006B.Size - 1)
                oMat1.LoadFromDataSource();
                functionReturnValue = true;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY006_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 대상자 조회
        /// </summary>
        private void PH_PY006_MTX01()
        {
            int i = 0;
            string sQry = string.Empty;

            string YM = string.Empty;
            string DocDate = string.Empty;

            string Param01 = string.Empty;
            string Param02 = string.Empty;
            string Param03 = string.Empty;
            double Param04 = 0;
            double Param05 = 0;
            string Param06 = string.Empty;

            short ErrNum = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", 100, false);

            //PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                Param01 = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                Param02 = oForm.Items.Item("JIGTYP").Specific.Value.ToString().Trim();
                Param03 = oForm.Items.Item("Gubun").Specific.Value.ToString().Trim();
                Param04 = Convert.ToDouble(oForm.Items.Item("UPHOBONG").Specific.Value);
                Param05 = Convert.ToDouble(oForm.Items.Item("UPAMT").Specific.Value);
                Param06 = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

                YM = oForm.Items.Item("YM").Specific.VALUE;
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(Param03))
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(Param06))
                {
                    ErrNum = 2;
                    throw new Exception();
                }

                sQry = "Select Count(*) From [@PH_PY006A] Where U_CLTCOD = '" + Param01 + "' and U_YM = '" + YM + "' and U_JIGTYP = '" + Param02 + "'";
                sQry = sQry + " and U_DocDate = '" + DocDate + "'";
                sQry = sQry + " and Canceled = 'N' ";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value > 0)
                {
                    ErrNum = 3;
                    throw new Exception();
                }

                sQry = "EXEC [PH_PY006_01] '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', " + Param04 + ", " + Param05 + ", '" + Param06 + "'";
                oRecordSet.DoQuery(sQry);

                oMat1.Clear();
                oMat1.FlushToDataSource();
                oMat1.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    ErrNum = 4;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PH_PY006B.InsertRecord((i));
                    }
                    oDS_PH_PY006B.Offset = i;
                    oDS_PH_PY006B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY006B.SetValue("U_TeamCode", i, oRecordSet.Fields.Item(0).Value);
                    oDS_PH_PY006B.SetValue("U_RspCode", i, oRecordSet.Fields.Item(1).Value);
                    oDS_PH_PY006B.SetValue("U_ClsCode", i, oRecordSet.Fields.Item(2).Value);
                    oDS_PH_PY006B.SetValue("U_MSTCOD", i, oRecordSet.Fields.Item(3).Value);
                    oDS_PH_PY006B.SetValue("U_FULLNAME", i, oRecordSet.Fields.Item(4).Value);
                    oDS_PH_PY006B.SetValue("U_JIGCOD", i, oRecordSet.Fields.Item(5).Value);
                    oDS_PH_PY006B.SetValue("U_GrpDat", i, oRecordSet.Fields.Item(6).Value);
                    oDS_PH_PY006B.SetValue("U_birthDat", i, oRecordSet.Fields.Item(7).Value);
                    oDS_PH_PY006B.SetValue("U_HOBYMM", i, oRecordSet.Fields.Item(8).Value);
                    oDS_PH_PY006B.SetValue("U_HOBONG", i, oRecordSet.Fields.Item(9).Value);
                    oDS_PH_PY006B.SetValue("U_HOBNAM", i, oRecordSet.Fields.Item(10).Value);
                    oDS_PH_PY006B.SetValue("U_STDAMT", i, oRecordSet.Fields.Item(11).Value);
                    oDS_PH_PY006B.SetValue("U_BNSAMT", i, oRecordSet.Fields.Item(12).Value);
                    oDS_PH_PY006B.SetValue("U_UPHOBONG", i, oRecordSet.Fields.Item(13).Value);
                    oDS_PH_PY006B.SetValue("U_UPHOBNAM", i, oRecordSet.Fields.Item(14).Value);
                    oDS_PH_PY006B.SetValue("U_UPSTDAMT", i, oRecordSet.Fields.Item(15).Value);
                    oDS_PH_PY006B.SetValue("U_UPBNSAMT", i, oRecordSet.Fields.Item(16).Value);
                    oDS_PH_PY006B.SetValue("U_APPLYYN", i, oRecordSet.Fields.Item(17).Value);
                    oDS_PH_PY006B.SetValue("U_PeakYN", i, oRecordSet.Fields.Item(18).Value);
                    oDS_PH_PY006B.SetValue("U_LineMemo", i, oRecordSet.Fields.Item(19).Value);
                    oRecordSet.MoveNext();
                    ProgressBar01.Value = ProgressBar01.Value + 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }
                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();
                oForm.Update();
            }
            catch(Exception ex)
            {
                ProgressBar01.Stop(); //StatusBar를 ProgressBar01가 점유하고 있기 때문에 오류 메시지를 출력하기 위해 ProgressBar01 정지

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("승호기준은 필수입니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    //dataHelpClass.MDC_GF_Message("승호기준은 필수입니다. 확인바랍니다.", "E");
                    oForm.Items.Item("Gubun").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("승호일자는 필수입니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    //dataHelpClass.MDC_GF_Message("승호일자는 필수입니다. 확인바랍니다.", "E");
                    oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("승호작업을 이미 등록했습니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    //dataHelpClass.MDC_GF_Message("승호작업을 이미 등록했습니다. 확인바랍니다.", "E");
                }
                else if (ErrNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    //dataHelpClass.MDC_GF_Message("결과가 존재하지 않습니다.", "E");
                }
                else
                {   
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY006_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 승호처리 및 발령사항 추가
        /// </summary>
        private void PH_PY006_MTX02()
        {

            //int i = 0;
            string sQry = string.Empty;

            string Param01 = string.Empty;
            string Param02 = string.Empty;
            string Param03 = string.Empty;
            string Param04 = string.Empty;

            string CLTCOD = string.Empty;
            string DocDate = string.Empty;
            string appNum = string.Empty;
            string JIGTYP = string.Empty;

            short ErrNum = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                Param01 = oForm.Items.Item("DocEntry").Specific.Value;
                Param02 = oForm.Items.Item("appNum").Specific.Value;
                Param03 = oForm.Items.Item("Canceled").Specific.Value;
                Param04 = PSH_Globals.oCompany.UserSignature.ToString();

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                appNum = oForm.Items.Item("appNum").Specific.Value;
                JIGTYP = oForm.Items.Item("JIGTYP").Specific.Value;

                sQry = "Select Count(*) From [@PH_PY006A] a Inner Join [@PH_PY006B] b On a.DocEntry = b.DocEntry ";
                sQry = sQry + " Inner Join [@PH_PY001G] c On b.U_MSTCOD = c.Code ";
                sQry = sQry + " Where a.DocEntry = '" + Param01 + "' and c.U_appNum = a.U_appNum and c.U_appType = '08' ";
                sQry = sQry + " and c.U_appDate = '" + DocDate + "'";

                oRecordSet.DoQuery(sQry);
                if (oRecordSet.Fields.Item(0).Value > 0)
                {   
                    ErrNum = 1;
                    throw new Exception();
                }

                sQry = "EXEC [PH_PY006_02] '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "'";
                oRecordSet.DoQuery(sQry);

                sQry = "EXEC [PH_PY006_03] '" + Param01 + "', '" + Param03 + "', '" + Param04 + "'";
                oRecordSet.DoQuery(sQry);

                PSH_Globals.SBO_Application.StatusBar.SetText("승호작업을 정상 처리했습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                //dataHelpClass.MDC_GF_Message("승호작업을 정상 처리했습니다.", "S");
            }
            catch(Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("승호작업을 이미 등록했습니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    //dataHelpClass.MDC_GF_Message("승호작업을 이미 처리했습니다. 확인바랍니다.", "E");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY006_MTX02_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 승호취소 및 발령관리 삭제
        /// </summary>
        private void PH_PY006_MTX03()
        {
            //int i = 0;
            string sQry = null;
            string DocDate = null;

            string Param01 = null;
            string Param02 = null;
            string Param03 = null;

            short ErrNum = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                Param01 = oForm.Items.Item("DocEntry").Specific.Value;
                Param02 = oForm.Items.Item("appNum").Specific.Value;
                Param03 = oForm.Items.Item("Canceled").Specific.Value;

                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

                sQry = "Select Count(*) From [@PH_PY006A] a Inner Join [@PH_PY006B] b On a.DocEntry = b.DocEntry ";
                sQry = sQry + " Inner Join [@PH_PY001G] c On b.U_MSTCOD = c.Code ";
                sQry = sQry + " Where a.DocEntry = '" + Param01 + "' and c.U_appNum = a.U_appNum and c.U_appType = '08' ";
                sQry = sQry + " and c.U_appDate = '" + DocDate + "'";

                oRecordSet.DoQuery(sQry);
                if (oRecordSet.Fields.Item(0).Value <= 0)
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                sQry = "EXEC [PH_PY006_02] '" + Param01 + "', '" + Param02 + "', '" + Param03 + "'";
                oRecordSet.DoQuery(sQry);

                sQry = "EXEC [PH_PY006_03] '" + Param01 + "', '" + Param03 + "'";
                oRecordSet.DoQuery(sQry);

                PSH_Globals.SBO_Application.StatusBar.SetText("승호작업을 취소 처리했습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                //dataHelpClass.MDC_GF_Message("승호작업을 취소 처리했습니다.", "S");
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("승호취소할 자료가 대상이 없거나 이미 취소했습니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    //dataHelpClass.MDC_GF_Message("승호취소 할 자료가 대상이 없거나 이미 취소했습니다. 확인바랍니다.", "E");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY006_MTX03_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY006_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY006A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return functionReturnValue;
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY006_Validate_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }

            return functionReturnValue;
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

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
                            if (PH_PY006_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY006_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        //대상자조회
                        PH_PY006_MTX01();
                    }
                    else if (pVal.ItemUID == "Btn02")
                    {
                        //승호처리 및 발령사항 추가
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (oForm.Items.Item("Canceled").Specific.Value == "N")
                            {
                                if (string.IsNullOrEmpty(oForm.Items.Item("appNum").Specific.Value.ToString().Trim()))
                                {
                                    //PSH_Globals.SBO_Application.SetStatusBarMessage("발령번호가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                    PSH_Globals.SBO_Application.StatusBar.SetText("발령번호가 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                }
                                else
                                {
                                    PH_PY006_MTX02();
                                }
                            }
                            else
                            {
                                //PSH_Globals.SBO_Application.SetStatusBarMessage("취소된 문서는 처리할 수 없습니다..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                PSH_Globals.SBO_Application.StatusBar.SetText("취소된 문서는 처리할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            }
                        }
                        else
                        {
                            //PSH_Globals.SBO_Application.SetStatusBarMessage("추가(수정)후 처리바랍니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            PSH_Globals.SBO_Application.StatusBar.SetText("추가(수정)후 처리바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        }
                    }
                    else if (pVal.ItemUID == "Btn03")
                    {
                        //승호취소 처리
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (oForm.Items.Item("Canceled").Specific.VALUE == "Y")
                            {
                                PH_PY006_MTX03();
                            }
                            else
                            {
                                //PSH_Globals.SBO_Application.SetStatusBarMessage("취소처리된 문서만 취소처리할 수 있습니다..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                PSH_Globals.SBO_Application.StatusBar.SetText("취소처리된 문서만 취소처리할 수 있습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            }
                        }
                        else
                        {
                            //PSH_Globals.SBO_Application.SetStatusBarMessage("추가(갱신)후 처리바랍니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            PSH_Globals.SBO_Application.StatusBar.SetText("추가(갱신)후 처리바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                                PH_PY006_FormItemEnabled();
                                PH_PY006_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY006_FormItemEnabled();
                                PH_PY006_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY006_FormItemEnabled();
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
                        //PH_PY006_AddMatrixRow

                        if (pVal.ItemUID == "JIGTYP" || pVal.ItemUID == "CLTCOD")
                        {
                            oMat1.Clear();
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (string.IsNullOrEmpty(pVal.ColUID))
                            {

                                oMat1.FlushToDataSource();
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
                                oLastItemUID = pVal.ItemUID;
                                oLastColUID = pVal.ColUID;
                                oLastColRow = pVal.Row;

                                oMat1.SelectRow(pVal.Row, true, false);
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
                                oDS_PH_PY006A.SetValue("U_CntcName", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.VALUE + "'", ""));
                                break;

                            case "Mat01":

                                if (string.IsNullOrEmpty(pVal.ColUID))
                                {
                                    oMat1.FlushToDataSource();
                                    oMat1.LoadFromDataSource();
                                    PH_PY006_AddMatrixRow();
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
                    PH_PY006_FormItemEnabled();
                    PH_PY006_AddMatrixRow();
                    oMat1.AutoResizeColumns();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY006A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY006B);
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
                    //if(pVal.ItemUID == "Code")
                    //{
                    //    dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY006A", "Code", "", 0, "", "", "");
                    //}
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
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY006A", "DocEntry"); //접속자 권한에 따른 사업장 보기
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY006_FormItemEnabled();
                            PH_PY006_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        //Case "1293":
                        //  Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
                        case "1281": //문서찾기
                            PH_PY006_FormItemEnabled();
                            PH_PY006_AddMatrixRow();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY006_FormItemEnabled();
                            PH_PY006_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY006_FormItemEnabled();
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
