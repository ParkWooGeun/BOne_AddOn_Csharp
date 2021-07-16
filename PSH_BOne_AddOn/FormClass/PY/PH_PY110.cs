using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 개인별상여율등록
    /// </summary>
    internal class PH_PY110 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY110A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY110B;
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY110.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY110_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY110");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";
                
                oForm.Freeze(true);
                PH_PY110_CreateItems();
                PH_PY110_EnableMenus();
                PH_PY110_SetDocument(oFormDocEntry);
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
        private void PH_PY110_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                oDS_PH_PY110A = oForm.DataSources.DBDataSources.Item("@PH_PY110A"); //헤더
                oDS_PH_PY110B = oForm.DataSources.DBDataSources.Item("@PH_PY110B"); //라인

                oMat01 = oForm.Items.Item("Mat1").Specific;

                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                //사업장
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //지급구분
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P212' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JOBGBN").Specific,"");
                oForm.Items.Item("JOBGBN").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("JOBGBN").DisplayDesc = true;

                //지급대상
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P213' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JOBTRG").Specific,"");
                oForm.Items.Item("JOBTRG").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("JOBTRG").DisplayDesc = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY110_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY110_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY110_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        private void PH_PY110_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PH_PY110_FormItemEnabled();
                    PH_PY110_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY110_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY110_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY110_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("JOBGBN").Enabled = true;
                    oForm.Items.Item("JOBTRG").Enabled = true;
                    oForm.Items.Item("Btn_Excel").Enabled = true;
                    oForm.Items.Item("Btn_Create").Enabled = true;
                    
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.Items.Item("YM").Specific.Value = DateTime.Now.ToString("yyyyMM"); //귀속년월

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("JOBGBN").Enabled = true;
                    oForm.Items.Item("JOBTRG").Enabled = true;
                    oForm.Items.Item("Btn_Excel").Enabled = false;
                    oForm.Items.Item("Btn_Create").Enabled = false;

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD",true); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("YM").Enabled = false;
                    oForm.Items.Item("JOBGBN").Enabled = false;
                    oForm.Items.Item("JOBTRG").Enabled = false;
                    oForm.Items.Item("Btn_Excel").Enabled = true;
                    oForm.Items.Item("Btn_Create").Enabled = true;
                    
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY110_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY110_FormClear
        /// </summary>
        private void PH_PY110_FormClear()
        {
            string DocEntry;

            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = DataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY110'", "");
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY110_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY110_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i;
            int k;
            string sCode;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //헤더
                if (string.IsNullOrEmpty(oDS_PH_PY110A.GetValue("U_CLTCOD", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY110A.GetValue("U_YM", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("귀속년월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY110A.GetValue("U_JOBGBN", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("지급구분은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("JOBGBN").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY110A.GetValue("U_JOBTRG", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("지급대상은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("JOBTRG").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                oMat01.FlushToDataSource();
                //라인
                if (oMat01.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        //구분
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("MSTCOD").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("사번은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat01.Columns.Item("MSTCOD").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }

                        //중복체크
                        for (k = i + 1; k <= oMat01.VisualRowCount - 2; k++)
                        {
                            if (oDS_PH_PY110B.GetValue("U_MSTCOD", i) == oDS_PH_PY110B.GetValue("U_MSTCOD", k).ToString().Trim())
                            {
                                PSH_Globals.SBO_Application.SetStatusBarMessage("내용이 중복입력되었습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                oMat01.Columns.Item("MSTCOD").Cells.Item(k + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                return functionReturnValue;
                            }
                        }
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return functionReturnValue;
                }

                sCode = oDS_PH_PY110A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY110A.GetValue("U_YM", 0).ToString().Trim() + oDS_PH_PY110A.GetValue("U_JOBGBN", 0).ToString().Trim() + oDS_PH_PY110A.GetValue("U_JOBTRG", 0).ToString().Trim();

                oDS_PH_PY110A.SetValue("Code", 0, sCode);
                oDS_PH_PY110A.SetValue("Name", 0, sCode);

                //데이터 중복 체크
                if (!string.IsNullOrEmpty(dataHelpClass.Get_ReData("Code", "Code", "[@PH_PY110A]", "'" + "sCode" + "'", "")))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("기존에 데이터가 존재합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return functionReturnValue;
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY110_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }

            return functionReturnValue;
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY110_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            short ErrNumm = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY110A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
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
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private void PH_PY110_AddMatrixRow()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);

                oMat01.FlushToDataSource();
                oRow = oMat01.VisualRowCount;

                if (oMat01.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY110B.GetValue("U_MSTNAM", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY110B.Size <= oMat01.VisualRowCount)
                        {
                            oDS_PH_PY110B.InsertRecord(oRow);
                        }
                        oDS_PH_PY110B.Offset = oRow;
                        oDS_PH_PY110B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY110B.SetValue("U_MSTCOD", oRow, "");
                        oDS_PH_PY110B.SetValue("U_MSTNAM", oRow, "");
                        oDS_PH_PY110B.SetValue("U_DPTCOD", oRow, "");
                        oDS_PH_PY110B.SetValue("U_DPTCOD", oRow, "");
                        oDS_PH_PY110B.SetValue("U_BNSRAT", oRow, "0");
                        oDS_PH_PY110B.SetValue("U_BNSAMT", oRow, "0");
                        oDS_PH_PY110B.SetValue("U_REMARK", oRow, "");
                        oMat01.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY110B.Offset = oRow - 1;
                        oDS_PH_PY110B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY110B.SetValue("U_MSTCOD", oRow - 1, "");
                        oDS_PH_PY110B.SetValue("U_MSTNAM", oRow - 1, "");
                        oDS_PH_PY110B.SetValue("U_DPTCOD", oRow - 1, "");
                        oDS_PH_PY110B.SetValue("U_DPTCOD", oRow - 1, "");
                        oDS_PH_PY110B.SetValue("U_BNSRAT", oRow - 1, "0");
                        oDS_PH_PY110B.SetValue("U_BNSAMT", oRow - 1, "0");
                        oDS_PH_PY110B.SetValue("U_REMARK", oRow - 1, "");
                        oMat01.LoadFromDataSource();
                    }
                }
                else if (oMat01.VisualRowCount == 0)
                {
                    oDS_PH_PY110B.Offset = oRow;
                    oDS_PH_PY110B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY110B.SetValue("U_MSTCOD", oRow, "");
                    oDS_PH_PY110B.SetValue("U_MSTNAM", oRow, "");
                    oDS_PH_PY110B.SetValue("U_DPTCOD", oRow, "");
                    oDS_PH_PY110B.SetValue("U_DPTCOD", oRow, "");
                    oDS_PH_PY110B.SetValue("U_BNSRAT", oRow, "0");
                    oDS_PH_PY110B.SetValue("U_BNSAMT", oRow, "0");
                    oDS_PH_PY110B.SetValue("U_REMARK", oRow, "");
                    oMat01.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY110_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                oForm.Freeze(false);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 엑셀 업로드
        /// </summary>
        [STAThread]
        private void PH_PY110_Excel_Upload()
        {
            int loopCount;
            int j;
            int CheckLine;
            int i;
            bool sucessFlag = false;
            short columnCount = 5; //엑셀 컬럼수
            short columnCount2 = 5; //엑셀 컬럼수
            string sFile;
            string sQry;
            double TOTCNT;
            int V_StatusCnt;
            int oProValue;
            int tRow;
            bool CheckYN;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            CommonOpenFileDialog commonOpenFileDialog = new CommonOpenFileDialog();

            commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("Excel Files", "*.xls;*.xlsx"));
            commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("모든 파일", "*.*"));

            if (commonOpenFileDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                sFile = commonOpenFileDialog.FileName;
            }
            else //Cancel 버튼 클릭
            {
                return;
            }

            if (string.IsNullOrEmpty(sFile))
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("파일을 선택해 주세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }
            
            //엑셀 Object 연결
            //암시적 객체참조 시 Excel.exe 메모리 반환이 안됨, 아래와 같이 명시적 참조로 선언
            Microsoft.Office.Interop.Excel.ApplicationClass xlapp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            Microsoft.Office.Interop.Excel.Workbooks xlwbs = xlapp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook xlwb = xlwbs.Open(sFile);
            Microsoft.Office.Interop.Excel.Sheets xlshs = xlwb.Worksheets;
            Microsoft.Office.Interop.Excel.Worksheet xlsh = (Microsoft.Office.Interop.Excel.Worksheet)xlshs[1];
            Microsoft.Office.Interop.Excel.Range xlCell = xlsh.Cells;
            Microsoft.Office.Interop.Excel.Range xlRange = xlsh.UsedRange;
            Microsoft.Office.Interop.Excel.Range xlRow = xlRange.Rows;

            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            oForm.Freeze(true);

            oMat01.Clear();
            oMat01.FlushToDataSource();
            oMat01.LoadFromDataSource();
            try
            {
                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("시작!", xlRow.Count, false);

                Microsoft.Office.Interop.Excel.Range[] t = new Microsoft.Office.Interop.Excel.Range[columnCount2 + 1];
                for (loopCount = 1; loopCount <= columnCount2; loopCount++)
                {
                    t[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[1, loopCount];
                }

                // 첫 타이틀 비교
                if (Convert.ToString(t[1].Value) != "사번")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("A열 첫번째 행 타이틀은 사번", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[2].Value) != "성명")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("B열 두번째 행 타이틀은 성명", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[3].Value) != "적용률")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("C열 세번째 행 타이틀은 적용률", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[4].Value) != "적용금액")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("D열 세번째 행 타이틀은 적용금액", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[5].Value) != "비고")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("E열 세번째 행 타이틀은 비고", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }

                //프로그레스 바    ///////////////////////////////////////
                ProgressBar01.Text = "데이터 읽는중...!";

                //최대값 구하기 ///////////////////////////////////////
                TOTCNT = xlsh.UsedRange.Rows.Count;

                V_StatusCnt = Convert.ToInt32(Math.Round(TOTCNT / 50, 0));
                oProValue = 1;
                tRow = 1;
                ///////////////////////////////////////////////////////

                CheckLine = 0;
                for (i = 2; i <= xlsh.UsedRange.Rows.Count; i++)
                {
                    Microsoft.Office.Interop.Excel.Range[] r = new Microsoft.Office.Interop.Excel.Range[columnCount + 1];

                    for (loopCount = 1; loopCount <= columnCount; loopCount++)
                    {
                        r[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[i, loopCount];
                    }

                    //사번 존재 여부 체크
                    if (dataHelpClass.Value_ChkYn("[@PH_PY001A]", "Code", "'" + Convert.ToString(r[1].Value) + "'", "") == true)
                    {
                        PSH_Globals.SBO_Application.StatusBar.SetText(i + "행의 [" + Convert.ToString(r[1].Value) + " ] 사번이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        throw new Exception();
                    }
                    else
                    {
                        CheckYN = false;
                        for (j = 0; j <= oDS_PH_PY110B.Size - 1; j++)
                        {

                            if (Convert.ToString(r[1].Value) == oDS_PH_PY110B.GetValue("U_MSTCOD", j).ToString().Trim())
                            {
                                CheckYN = true;
                                CheckLine = j;
                            }
                        }

                        //마지막행 제거
                        if (string.IsNullOrEmpty(oDS_PH_PY110B.GetValue("U_MSTCOD", oDS_PH_PY110B.Size - 1).ToString().Trim()))
                        {
                            oDS_PH_PY110B.RemoveRecord((oDS_PH_PY110B.Size - 1));
                        }
                        //사원마스터에서 사번에 대한 정보 가져오기
                        sQry = "  select U_FullName, U_TeamCode, U_CodeNm ";
                        sQry += " FROM [@PH_PY001A] T0 INNER JOIN [@PS_HR200L] T1 ON T0.U_TeamCode = T1.U_Code";
                        sQry += " WHERE T0.Code='" + Convert.ToString(r[1].Value) + "'";
                        oRecordSet01.DoQuery(sQry);

                        //Mat1에 업로드엑셀 사번이 존재 유무 CheckYN
                        if (CheckYN == true)
                        {
                            oDS_PH_PY110B.SetValue("U_MSTNAM", CheckLine, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                            oDS_PH_PY110B.SetValue("U_DPTCOD", CheckLine, oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                            oDS_PH_PY110B.SetValue("U_DPTNAM", CheckLine, oRecordSet01.Fields.Item(2).Value.ToString().Trim());
                            oDS_PH_PY110B.SetValue("U_BNSRAT", CheckLine, Convert.ToString(r[3].Value));
                            oDS_PH_PY110B.SetValue("U_BNSAMT", CheckLine, Convert.ToString(r[4].Value));
                            oDS_PH_PY110B.SetValue("U_REMARK", CheckLine, Convert.ToString(r[5].Value));
                        }
                        else
                        {
                            oDS_PH_PY110B.InsertRecord(oDS_PH_PY110B.Size);
                            oDS_PH_PY110B.Offset = oDS_PH_PY110B.Size - 1;
                            oDS_PH_PY110B.SetValue("U_MSTCOD", oDS_PH_PY110B.Size - 1, Convert.ToString(r[1].Value));
                            oDS_PH_PY110B.SetValue("U_MSTNAM", oDS_PH_PY110B.Size - 1, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                            oDS_PH_PY110B.SetValue("U_DPTCOD", oDS_PH_PY110B.Size - 1, oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                            oDS_PH_PY110B.SetValue("U_DPTNAM", oDS_PH_PY110B.Size - 1, oRecordSet01.Fields.Item(2).Value.ToString().Trim());
                            oDS_PH_PY110B.SetValue("U_BNSRAT", oDS_PH_PY110B.Size - 1, Convert.ToString(r[3].Value));
                            oDS_PH_PY110B.SetValue("U_BNSAMT", oDS_PH_PY110B.Size - 1, Convert.ToString(r[4].Value));
                            oDS_PH_PY110B.SetValue("U_REMARK", oDS_PH_PY110B.Size - 1, Convert.ToString(r[5].Value));
                        }

                        if ((TOTCNT > 50 && tRow == oProValue * V_StatusCnt) || TOTCNT <= 50)
                        {
                            ProgressBar01.Text = tRow + "/ " + TOTCNT + " 건 처리중...!";
                            ProgressBar01.Value += 1;
                        }
                        tRow += 1;
                    }
                }
                ProgressBar01.Value += 1;
                ProgressBar01.Text = ProgressBar01.Value + "/" + (xlRow.Count - 1) + "건 Loding...!";

                //라인번호 재정의
                for (i = 0; i <= oDS_PH_PY110B.Size - 1; i++)
                {
                    oDS_PH_PY110B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();

                sucessFlag = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY110_Excel_Upload:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                xlapp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRow);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCell);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsh);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlshs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwbs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp);
                oForm.Freeze(false);
                ProgressBar01.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);

                if (sucessFlag == true)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("엑셀 Loding 완료", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
        }

        /// <summary>
        /// MSTCOD_Setting
        /// </summary>
        private void MSTCOD_Setting()
        {
            string sQry;
            int iRow;
            int MatRowCnt;
            int oRow;
            short ErrNum = 0;
            string CLTCOD;            ////사업장
            string JOBYMM;            ////귀속년월
            string JOBTRG;            ////지급대상
            double TOTCNT;
            int V_StatusCnt;
            int oProValue;
            int tRow;
            string CheckValue;

            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                CLTCOD = oDS_PH_PY110A.GetValue("U_CLTCOD", 0).ToString().Trim();
                JOBYMM = oDS_PH_PY110A.GetValue("U_YM", 0).ToString().Trim();
                JOBTRG = oDS_PH_PY110A.GetValue("U_JOBTRG", 0).ToString().Trim();
                if (string.IsNullOrEmpty(JOBYMM))
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                sQry = "EXEC PH_PY110 '" + CLTCOD + "','" + JOBYMM + "','" + JOBTRG + "'";
                oRecordSet01.DoQuery(sQry);

                if (oRecordSet01.RecordCount == 0)
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                //최대값 구하기 ///////////////////////////////////////
                TOTCNT = oRecordSet01.RecordCount / 50;

                V_StatusCnt = Convert.ToInt32(Math.Round(TOTCNT, 0));
                oProValue = 1;
                tRow = 1;
                ///////////////////////////////////////////////////////

                oDS_PH_PY110B.RemoveRecord((oDS_PH_PY110B.Size - 1));

                MatRowCnt = oMat01.VisualRowCount - 2;
                oRow = oMat01.VisualRowCount - 1;

                while (!oRecordSet01.EoF)
                {
                    CheckValue = "N";
                    for (iRow = 0; iRow <= MatRowCnt; iRow++)
                    {
                        if (oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim() == oDS_PH_PY110B.GetValue("U_MSTCOD", iRow).ToString().Trim())
                        {
                            CheckValue = "Y";
                            break;
                        }
                    }

                    if (CheckValue == "N")
                    {
                        oDS_PH_PY110B.InsertRecord(oRow);
                        oDS_PH_PY110B.Offset = oRow;
                        oDS_PH_PY110B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY110B.SetValue("U_MSTCOD", oRow, oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim());    //사원번호
                        oDS_PH_PY110B.SetValue("U_MSTNAM", oRow, oRecordSet01.Fields.Item("MSTNAM").Value.ToString().Trim());    //성명
                        oDS_PH_PY110B.SetValue("U_DPTCOD", oRow, oRecordSet01.Fields.Item("DPTCOD").Value.ToString().Trim());    //부서
                        oDS_PH_PY110B.SetValue("U_DPTNAM", oRow, oRecordSet01.Fields.Item("DPTNAM").Value.ToString().Trim());    //부서명
                        oDS_PH_PY110B.SetValue("U_BNSRAT", oRow, "100");                                                         //적용율(%)
                        oDS_PH_PY110B.SetValue("U_BNSAMT", oRow, "0");                                                           //적용금액
                        oDS_PH_PY110B.SetValue("U_REMARK", oRow, "");
                        oRow += 1;
                    }

                    tRow += 1;
                    if ((TOTCNT > 50 && tRow == oProValue * V_StatusCnt) || TOTCNT <= 50)
                    {
                        ProgressBar01.Text = tRow + "/ " + TOTCNT + " 건 처리중...!";
                        oProValue += 1;
                        ProgressBar01.Value = oProValue;
                    }

                    oRecordSet01.MoveNext();
                }

                PSH_Globals.SBO_Application.StatusBar.SetText("작업이 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                if (oRow == MatRowCnt + 1)
                {
                    ErrNum = 3;
                    throw new Exception();
                }
                oMat01.LoadFromDataSource();
                PH_PY110_AddMatrixRow();
            }
            catch (Exception ex)
            {
                ProgressBar01.Stop();
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("귀속년월은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("생성할 대상자가 없습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("이미 모든 대상자가 생성되어 있어 추가생성된 대상자가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("MSTCOD_Setting Error : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY110_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                    }
                    if (pVal.ItemUID == "Btn_Excel")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY110_Excel_Upload);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();

                        PH_PY110_AddMatrixRow();
                    }
                    if (pVal.ItemUID == "Btn_Create")
                    {
                        MSTCOD_Setting();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (pVal.ActionSuccess == true)
                        {
                            PH_PY110_FormItemEnabled();
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
                if (pVal.BeforeAction == true && pVal.ItemUID == "Mat1" && pVal.ColUID == "MSTCOD" && pVal.CharPressed == 9)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value))
                    {
                        oMat01.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                        BubbleEvent = false;
                    }
                    else
                    {
                        if (dataHelpClass.Value_ChkYn("[@PH_PY001A]", "Code",  "'" + oMat01.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value + "'", "") == true)
                        {
                            oMat01.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
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
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
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
                        if (pVal.ItemUID == "Mat1" && pVal.ColUID == "MSTCOD")
                        {
                            if (!string.IsNullOrEmpty(oMat01.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value))
                            {
                                oMat01.Columns.Item("MSTNAM").Cells.Item(pVal.Row).Specific.Value = dataHelpClass.Get_ReData("U_FULLNAME", "Code", "[@PH_PY001A]", "'" + oMat01.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value + "'","");
                                oMat01.Columns.Item("DPTCOD").Cells.Item(pVal.Row).Specific.Value = dataHelpClass.Get_ReData("U_TeamCode", "Code", "[@PH_PY001A]", "'" + oMat01.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value + "'", "");
                                oMat01.Columns.Item("DPTNAM").Cells.Item(pVal.Row).Specific.Value = dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oMat01.Columns.Item("DPTCOD").Cells.Item(pVal.Row).Specific.Value + "'", " AND Code = '1'");
                            }
                            PH_PY110_AddMatrixRow();
                            oMat01.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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

                    PH_PY110_FormItemEnabled();
                    PH_PY110_AddMatrixRow();
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
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Mat1":
                            if (pVal.Row > 0)
                            {
                                oMat01.SelectRow(pVal.Row, true, false);

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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY110A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY110B);
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
                            dataHelpClass.AuthorityCheck( oForm,  "CLTCOD",  "@PH_PY110A",  "Code"); //접속자 권한에 따른 사업장 보기
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY110_FormItemEnabled();
                            PH_PY110_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY110_FormItemEnabled();
                            PH_PY110_AddMatrixRow();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY110_FormItemEnabled();
                            PH_PY110_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY110_FormItemEnabled();
                            break;
                        case "1293":
                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                oMat01.FlushToDataSource();

                                while ((i <= oDS_PH_PY110B.Size - 1))
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY110B.GetValue("U_MSTCOD", i)))
                                    {
                                        oDS_PH_PY110B.RemoveRecord(i);
                                        i = 0;
                                    }
                                    else
                                    {
                                        i += 1;
                                    }
                                }

                                for (i = 0; i <= oDS_PH_PY110B.Size; i++)
                                {
                                    oDS_PH_PY110B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                }

                                oMat01.LoadFromDataSource();
                            }
                            PH_PY110_AddMatrixRow();
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

