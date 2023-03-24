using System;
using System.IO;
using SAPbouiCOM;
using System.Collections.Generic;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;
using MsOutlook = Microsoft.Office.Interop.Outlook;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 건강보험 연말정산
    /// </summary>
    internal class PH_PY900 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY900A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY900B;
        private string oLastItemUID;     //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow;         //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private bool CheckDataApply; //적용버튼 실행여부
        private string CLTCOD; //사업장
        private string YY; //적용연월

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY900.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY900_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY900");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";
                
                oForm.Freeze(true);
                PH_PY900_CreateItems();
                PH_PY900_EnableMenus();
                PH_PY900_SetDocument(oFormDocEntry);
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
        private void PH_PY900_CreateItems()
        {
            try
            {
                oForm.Freeze(true);
                oDS_PH_PY900A = oForm.DataSources.DBDataSources.Item("@PH_PY900A");
                oDS_PH_PY900B = oForm.DataSources.DBDataSources.Item("@PH_PY900B");

                oMat1 = oForm.Items.Item("Mat1").Specific;
                
                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                CheckDataApply = false;

                oForm.Items.Item("CLTCOD").DisplayDesc = true; //사업장
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY900_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY900_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY900_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        private void PH_PY900_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PH_PY900_FormItemEnabled();
                    PH_PY900_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY900_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY900_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY900_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YY").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = false;
                    oForm.Items.Item("Mat1").Enabled = true;
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YY").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = false;
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("YY").Enabled = false;
                    oForm.Items.Item("Comments").Enabled = false;
                    oForm.Items.Item("Mat1").Enabled = false;
                    }
                    else
                    {
                        oForm.Items.Item("FieldCo").Enabled = true;
                        oForm.Items.Item("Mat1").Enabled = true;
                    }
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY900_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        private void PH_PY900_AddMatrixRow()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY900B.GetValue("U_LineNum", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY900B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY900B.InsertRecord(oRow);
                        }
                        oDS_PH_PY900B.Offset = oRow;
                        oDS_PH_PY900B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY900B.SetValue("U_govID", oRow, "");
                        oDS_PH_PY900B.SetValue("U_MSTNAM", oRow, "");
                        oDS_PH_PY900B.SetValue("U_BuSum", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_BuGun", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_BuJang", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_BoSuTo", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_WorkM", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_BoSuM", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_HwakSum", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_HwakGun", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_HwakJang", oRow, "0");
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY900B.Offset = oRow - 1;
                        oDS_PH_PY900B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY900B.SetValue("U_govID", oRow, "");
                        oDS_PH_PY900B.SetValue("U_MSTNAM", oRow, "");
                        oDS_PH_PY900B.SetValue("U_BuSum", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_BuGun", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_BuJang", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_BoSuTo", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_WorkM", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_BoSuM", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_HwakSum", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_HwakGun", oRow, "0");
                        oDS_PH_PY900B.SetValue("U_HwakJang", oRow, "0");
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY900B.Offset = oRow;
                    oDS_PH_PY900B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY900B.SetValue("U_govID", oRow, "");
                    oDS_PH_PY900B.SetValue("U_MSTNAM", oRow, "");
                    oDS_PH_PY900B.SetValue("U_BuSum", oRow, "0");
                    oDS_PH_PY900B.SetValue("U_BuGun", oRow, "0");
                    oDS_PH_PY900B.SetValue("U_BuJang", oRow, "0");
                    oDS_PH_PY900B.SetValue("U_BoSuTo", oRow, "0");
                    oDS_PH_PY900B.SetValue("U_WorkM", oRow, "0");
                    oDS_PH_PY900B.SetValue("U_BoSuM", oRow, "0");
                    oDS_PH_PY900B.SetValue("U_HwakSum", oRow, "0");
                    oDS_PH_PY900B.SetValue("U_HwakGun", oRow, "0");
                    oDS_PH_PY900B.SetValue("U_HwakJang", oRow, "0");
                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY900_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        private bool PH_PY900_DataValidCheck()
        {
            bool returnValue = false;
            int i;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oDS_PH_PY900A.GetValue("U_CLTCOD", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //적용시작월
                if (string.IsNullOrEmpty(oDS_PH_PY900A.GetValue("U_YY", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("적용시작월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("YY").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //Code & Name 생성
                oDS_PH_PY900A.SetValue("Code", 0, oDS_PH_PY900A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY900A.GetValue("U_YY", 0).ToString().Trim());
                oDS_PH_PY900A.SetValue("NAME", 0, oDS_PH_PY900A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY900A.GetValue("U_YY", 0).ToString().Trim());

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (!string.IsNullOrEmpty(dataHelpClass.Get_ReData("Code", "Code", "[@PH_PY900A]", "'" + oDS_PH_PY900A.GetValue("Code", 0).ToString().Trim() + "'", "")))
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("이미 존재하는 코드입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        return returnValue;
                    }
                }

                //라인
                if (oMat1.VisualRowCount >= 1)
                {
                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                    {
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return returnValue;
                }

                oMat1.FlushToDataSource();

                //Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY900B.Size > 1)
                {
                    oDS_PH_PY900B.RemoveRecord(oDS_PH_PY900B.Size - 1);
                }

                oMat1.LoadFromDataSource();
                returnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY900_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// PH_PY900_FormClear
        /// </summary>
        private void PH_PY900_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = DataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY900'", "");
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY900_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY900_Validate(string ValidateType)
        {
            bool returnValue = false;
            short ErrNumm = 0;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY900A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
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

                returnValue = true;
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

            return returnValue;
        }
        
        /// <summary>
        /// 엑셀 파일 업로드
        /// </summary>
        [STAThread]
        private void PH_PY900_Excel_Upload()
        {
            int loopCount;
            int j;
            int CheckLine;
            int i;
            bool sucessFlag = false;
            short columnCount = 29; //엑셀 컬럼수
            short columnCount2 = 29; //엑셀 컬럼수
            string sFile;
            double TOTCNT;
            int V_StatusCnt;
            int oProValue;
            int tRow;
            string FullName;
            string pgovID;
            string sQry;
            double bua;
            double hwaka;

            //bool CheckYN = true;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            CommonOpenFileDialog commonOpenFileDialog = new CommonOpenFileDialog();

            commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("Excel Files", "*.xls;*.xlsx"));
            commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("모든 파일", "*.*"));
            commonOpenFileDialog.IsFolderPicker = false;

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
                return;
            }
            else
            {
                oForm.Items.Item("Comments").Specific.Value = sFile;
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

            oMat1.Clear();
            oMat1.FlushToDataSource();
            oMat1.LoadFromDataSource();
            try
            {
                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("시작!", xlRow.Count, false);
                Microsoft.Office.Interop.Excel.Range[] t = new Microsoft.Office.Interop.Excel.Range[columnCount2 + 1];
                for (loopCount = 1; loopCount <= columnCount2; loopCount++)
                {
                    t[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[1, loopCount];
                }

                // 첫 타이틀 비교
                if (Convert.ToString(t[1].Value) != "연번")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("A열 첫번째 행 타이틀은 연번", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[7].Value) != "성명")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("G열 첫번째 행 타이틀은 성명", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[8].Value) != "주민등록번호")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("H열 두번째 행 타이틀은 성명", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[9].Value) != "부과한총보험료")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("D열 세번째 행 타이틀은 부과한총보험료", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[10].Value) != "건강")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("J열 세번째 행 타이틀은 건강", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[11].Value) != "장기요양")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("K열 세번째 행 타이틀은 장기요양", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[12].Value) != "연간보수총액")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("L열 세번째 행 타이틀은 연간보수총액", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[13].Value) != "근무월수")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("M열 세번째 행 타이틀은 근무월수", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }

                if (Convert.ToString(t[14].Value) != "보수월액")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("N열 세번째 행 타이틀은 보수월액", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[17].Value) != "확정보험료")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("Q열 세번째 행 타이틀은 확정보험료", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[18].Value) != "건강")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("R열 세번째 행 타이틀은 건강", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[19].Value) != "장기요양")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("S열 세번째 행 타이틀은 장기요양", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
             
                //프로그레스 바
                ProgressBar01.Text = "데이터 읽는중...!";

                //최대값 구하기
                TOTCNT = xlsh.UsedRange.Rows.Count;

                V_StatusCnt = Convert.ToInt32(Math.Round(TOTCNT / 50, 0));
                oProValue = 1;
                tRow = 1;
                
                for (i = 2; i <= xlsh.UsedRange.Rows.Count; i++)
                {
                    Microsoft.Office.Interop.Excel.Range[] r = new Microsoft.Office.Interop.Excel.Range[columnCount + 1];

                    for (loopCount = 1; loopCount <= columnCount; loopCount++)
                    {
                        r[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[i, loopCount];
                    }
                    for (j = 0; j <= oDS_PH_PY900B.Size - 1; j++)
                    {

                        if (Convert.ToString(r[1].Value) == oDS_PH_PY900B.GetValue("U_govID", j).ToString().Trim())
                        {
                            CheckLine = j;
                        }
                    }

                    //마지막행 제거
                    if (string.IsNullOrEmpty(oDS_PH_PY900B.GetValue("U_govID", oDS_PH_PY900B.Size - 1).ToString().Trim()))
                    {
                        oDS_PH_PY900B.RemoveRecord(oDS_PH_PY900B.Size - 1);
                    }

                    oDS_PH_PY900B.InsertRecord(oDS_PH_PY900B.Size);
                    oDS_PH_PY900B.Offset = oDS_PH_PY900B.Size - 1;
                    oDS_PH_PY900B.SetValue("U_LineNum", oDS_PH_PY900B.Size - 1, Convert.ToString(r[1].Value));
                    oDS_PH_PY900B.SetValue("U_MSTNAM", oDS_PH_PY900B.Size - 1, Convert.ToString(r[7].Value));
                    oDS_PH_PY900B.SetValue("U_govID", oDS_PH_PY900B.Size - 1, Convert.ToString(r[8].Value));
                    oDS_PH_PY900B.SetValue("U_BuSum", oDS_PH_PY900B.Size - 1, Convert.ToString(r[9].Value));
                    oDS_PH_PY900B.SetValue("U_BuGun", oDS_PH_PY900B.Size - 1, Convert.ToString(r[10].Value));
                    oDS_PH_PY900B.SetValue("U_BuJang", oDS_PH_PY900B.Size - 1, Convert.ToString(r[11].Value));
                    oDS_PH_PY900B.SetValue("U_BoSuTo", oDS_PH_PY900B.Size - 1, Convert.ToString(r[12].Value));
                    oDS_PH_PY900B.SetValue("U_WorkM", oDS_PH_PY900B.Size - 1, Convert.ToString(r[13].Value));
                    oDS_PH_PY900B.SetValue("U_BoSuM", oDS_PH_PY900B.Size - 1, Convert.ToString(r[14].Value));
                    oDS_PH_PY900B.SetValue("U_HwakSum", oDS_PH_PY900B.Size - 1, Convert.ToString(r[17].Value));
                    oDS_PH_PY900B.SetValue("U_HwakGun", oDS_PH_PY900B.Size - 1, Convert.ToString(r[18].Value));
                    oDS_PH_PY900B.SetValue("U_HwakJang", oDS_PH_PY900B.Size - 1, Convert.ToString(r[19].Value));
                 
                    if ((TOTCNT > 50 && tRow == oProValue * V_StatusCnt) || TOTCNT <= 50)
                    {
                        ProgressBar01.Text = tRow + "/ " + TOTCNT + " 건 처리중...!";
                        ProgressBar01.Value += 1;
                    }
                    tRow += 1;
                }
                ProgressBar01.Value += 1;
                ProgressBar01.Text = ProgressBar01.Value + "/" + (xlRow.Count - 1) + "건 Loding...!";

                //라인번호 재정의
                for (i = 0; i <= oDS_PH_PY900B.Size - 1; i++)
                {
                    oDS_PH_PY900B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                }
                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();
                PH_PY900_AddMatrixRow();

                for (loopCount = 1; loopCount <= oMat1.VisualRowCount - 1; loopCount++)
                {
                    FullName = oMat1.Columns.Item("MSTNAM").Cells.Item(loopCount).Specific.Value;
                    pgovID = codeHelpClass.Left(oMat1.Columns.Item("govID").Cells.Item(loopCount).Specific.Value, 6) + codeHelpClass.Right(oMat1.Columns.Item("govID").Cells.Item(loopCount).Specific.Value, 7);
                    CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();

                    bua = Convert.ToDouble(oMat1.Columns.Item("BuSum").Cells.Item(loopCount).Specific.Value);
                    hwaka =Convert.ToDouble(oMat1.Columns.Item("HwakSum").Cells.Item(loopCount).Specific.Value);

                    sQry = "Select Code, U_eMail From [@PH_PY001A] wHERE U_status <> '5' and U_CLTCOD = '" + CLTCOD + "' and U_FullName = '" + FullName + "' And U_govID = '" + pgovID + "'";
                    oRecordSet01.DoQuery(sQry);
                   
                    if (oRecordSet01.RecordCount > 0)
                    {
                        oMat1.Columns.Item("MSTCOD").Cells.Item(loopCount).Specific.Value = oRecordSet01.Fields.Item(0).Value;
                        oMat1.Columns.Item("eMail").Cells.Item(loopCount).Specific.Value = oRecordSet01.Fields.Item(1).Value;
                        oMat1.Columns.Item("govID").Cells.Item(loopCount).Specific.Value = codeHelpClass.Left(pgovID, 6) + "-" + pgovID.ToString().Substring(6, 1) + "******";
                        (oMat1.Columns.Item("JungSan").Cells.Item(loopCount).Specific.Value) = hwaka - bua;
                    }
                    else
                    {
                        oMat1.Columns.Item("govID").Cells.Item(loopCount).Specific.Value = codeHelpClass.Left(pgovID, 6) + "-" + pgovID.ToString().Substring(6, 1) + "******";
                    }
                }
                oMat1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY900_Excel_Upload:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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

                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                if (sucessFlag == true)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("엑셀 Loding 완료", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                oForm.Freeze(false);
            }
        }


        /// <summary>
        /// PDF만들기
        /// </summary>
        [STAThread]
        private bool Make_PDF_File(String p_MSTCOD)
        {
            bool ReturnValue = false;
            string WinTitle;
            string ReportName = String.Empty;
            string CLTCOD;
            string YY;
            string pMSTNAM;
            string pCode;
            string Main_Folder;
            string Sub_Folder1;
            string Sub_Folder2;
            string sQry1;
            string sQry;
            string ExportString;
            string psgovID;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim(); //사업장
                YY = oForm.Items.Item("YY").Specific.Value.ToString().Trim(); //년월

                WinTitle = "[PH_PY900] 건강보험연말정산";
                ReportName = "PH_PY900_01.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD));
                dataPackParameter.Add(new PSH_DataPackClass("@YY", YY));
                dataPackParameter.Add(new PSH_DataPackClass("@MSTCOD", p_MSTCOD));

                Main_Folder = @"C:\PSH_건강보험연말정산";
                Sub_Folder1 = @"C:\PSH_건강보험연말정산\" + YY + "";
                Sub_Folder2 = @"C:\PSH_건강보험연말정산\" + YY + @"\" + CLTCOD + "";

                Dir_Exists(Main_Folder);
                Dir_Exists(Sub_Folder1);
                Dir_Exists(Sub_Folder2);
                
                sQry1 = " exec [PH_PY900_01] '" + CLTCOD + "','" + YY + "','" + p_MSTCOD + "'";
                oRecordSet01.DoQuery(sQry1);

                pCode=  oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                pMSTNAM = oRecordSet01.Fields.Item(6).Value.ToString().Trim();


                ExportString = Sub_Folder2 + @"\" + YY + "_" +  pMSTNAM + ".pdf";

                sQry = "Select RIGHT(U_govID,7) From [@PH_PY001A]";
                sQry += "WHERE  Code ='" + p_MSTCOD + "'";
                oRecordSet01.DoQuery(sQry);
                psgovID = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, ExportString,100);

                // Open an existing document. Providing an unrequired password is ignored.
                PdfDocument document = PdfReader.Open(ExportString, PdfDocumentOpenMode.Modify);

                PdfSecuritySettings securitySettings = document.SecuritySettings;

                securitySettings.UserPassword = "manager";   //개개인암호
                securitySettings.OwnerPassword = psgovID;    //마스터암호

                // Restrict some rights.
                securitySettings.PermitAccessibilityExtractContent = false;
                securitySettings.PermitAnnotations = false;
                securitySettings.PermitAssembleDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitFormsFill = true;
                securitySettings.PermitFullQualityPrint = false;
                securitySettings.PermitModifyDocument = true;
                securitySettings.PermitPrint = false;

                // PDF문서 저장
                document.Save(ExportString);

                sQry = "Update [@PH_PY900B] Set U_SaveYN = 'Y' Where U_MSTCOD = '" + p_MSTCOD + "' And Code = '" + pCode + "'";
                oRecordSet01.DoQuery(sQry);

                ReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
            return ReturnValue;
        }

        /// <summary>
        /// 디렉토리 체크, 폴더 생성
        /// </summary>
        /// <param name="strDirName">경로</param>
        /// <returns></returns>
        private int Dir_Exists(string strDirName)
        {
            int ReturnValue = 0;

            try
            {
                DirectoryInfo di = new DirectoryInfo(strDirName); //DirectoryInfo 생성
                //DirectoryInfo.Exists로 폴더 존재유무 확인
                if (di.Exists)
                {
                    ReturnValue = 1;
                }
                else
                {
                    di.Create();
                    ReturnValue = 0;
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Make_PDF_File_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
            return ReturnValue;
        }

        /// <summary>
        /// Send_EMail
        /// </summary>
        /// <param name="p_MSTCOD"></param>
        /// <param name="p_Version"></param>
        /// <returns></returns>
        private bool Send_EMail(string p_MSTCOD)
        {
            bool ReturnValue = false;
            string strToAddress;
            string strSubject;
            string strBody;
            string Sub_Folder2;
            string sQry;
            string MSTNAM;
            string MSTCOD;
            string Version;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                MSTCOD = p_MSTCOD;
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim(); //사업장
                YY = oForm.Items.Item("YY").Specific.Value.ToString().Trim(); //년월
                Version = CLTCOD + YY;

                Sub_Folder2 = @"C:\PSH_건강보험연말정산\" + YY + @"\" + CLTCOD + "";

                sQry = "Select U_Subject, U_Remark From [@PH_PY900A] Where Code = '" + Version + "'";
                oRecordSet01.DoQuery(sQry);
                strSubject = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                strBody = oRecordSet01.Fields.Item(1).Value.ToString().Trim();

                sQry = "Select U_eMail, U_MSTNAM From [@PH_PY900B] Where U_MSTCOD = '" + MSTCOD + "' AND Code = '" + Version + "'";
                oRecordSet01.DoQuery(sQry);
                strToAddress = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                MSTNAM = oRecordSet01.Fields.Item(1).Value.ToString().Trim();

                //mail.From = new MailAddress("dakkorea1@gmail.com");
                MsOutlook.Application outlookApp = new MsOutlook.Application();
                if (outlookApp == null)
                {
                    throw new Exception();
                }
                MsOutlook.MailItem mail = (MsOutlook.MailItem)outlookApp.CreateItem(MsOutlook.OlItemType.olMailItem);

                mail.Subject = strSubject;
                mail.HTMLBody = strBody;
                mail.To = strToAddress;
                MsOutlook.Attachment oAttach = mail.Attachments.Add(Sub_Folder2 + @"\" + YY + "_" + MSTNAM +".pdf");
                mail.Send();

                mail = null;
                outlookApp = null;

                sQry = "Update [@PH_PY900B] Set U_SendYN = 'Y' Where U_MSTCOD = '" + MSTCOD + "' And Code = '" + Version + "'";
                oRecordSet01.DoQuery(sQry);

                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(Sub_Folder3 + @"\" + p_MSTCOD + "_개인별급여명세서_" + STDYER + "" + STDMON + ".pdf");

                //원래코드시작
                //SmtpClient smtp = new SmtpClient("smtp.naver.com");
                //SmtpClient smtp = new SmtpClient("pscsn.poongsan.co.kr");
                //SmtpClient smtp = new SmtpClient("smtp.office365.com");
                //SmtpClient smtp = new SmtpClient("smtp.gmail.com");

                //smtp.Port = 587; //네이버
                //smtp.Port = 25; //풍산
                //smtp.UseDefaultCredentials = true;
                //smtp.EnableSsl = true;
                //smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                //smtp.Timeout = 20000;

                //smtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;  //Naver 인 경우
                //smtp.Credentials = new NetworkCredential("2220501", "p2220501!"); //address, PW
                //smtp.Credentials = new NetworkCredential("wgpark@poongsan.co.kr", "1q2w3e4r)*"); //address, PW
                //smtp.Credentials = new NetworkCredential("dakkorea1@gmail.com", "dak440310*"); //address, PW

                //smtp.Send(mail);
                //원래코드 끝

                ReturnValue = true;
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Send_EMail_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            return ReturnValue;
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
            string p_MSTCOD;
            string MCLTCOD;
            string MYY;
            string sQry;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY900_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                    }

                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (pVal.ActionSuccess == true)
                        {
                            if (CheckDataApply == true)
                            {
                                CheckDataApply = false;
                            }
                            PH_PY900_FormItemEnabled();
                            PH_PY900_AddMatrixRow();
                        }
                    }
                    if (pVal.ItemUID == "Btn_UPLOAD")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY900_Excel_Upload);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                        PH_PY900_AddMatrixRow();
                    }

                    if (pVal.ItemUID == "Btn_Print")
                    {
                        MCLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim(); //사업장
                        MYY = oForm.Items.Item("YY").Specific.Value.ToString().Trim(); //년월

                        sQry = " SELECT Count(*) FROM [@PH_PY900A] WHERE U_CLTCOD ='" + MCLTCOD + "' AND U_YY ='" + MYY + "'";
                        oRecordSet01.DoQuery(sQry);

                        if(Convert.ToDouble(oRecordSet01.Fields.Item(0).Value) == 0)
                        {
                            PSH_Globals.SBO_Application.MessageBox("추가를 먼저하고 PDF저장을 누르세요.");
                            BubbleEvent = false;
                            return;
                        }

                        ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("PDF 파일 생성 시작!", 50, false);

                        for (int i = 0; i <= oMat1.VisualRowCount - 1; i++)
                        {
                            if (!string.IsNullOrEmpty(oDS_PH_PY900B.GetValue("U_MSTCOD", i).ToString().Trim()))
                            {
                                if (!string.IsNullOrEmpty(oDS_PH_PY900B.GetValue("U_eMail", i).ToString().Trim()))
                                {
                                    p_MSTCOD = oDS_PH_PY900B.GetValue("U_MSTCOD", i).ToString().Trim();
                                    if (Make_PDF_File(p_MSTCOD) == false)
                                    {
                                        errMessage = "PDF저장이 완료되지 않았습니다.";
                                        throw new Exception();
                                    }
                                }
                            }
                            ProgressBar01.Value += 1;
                            ProgressBar01.Text = ProgressBar01.Value + "/" + (oMat1.VisualRowCount) + "건 PDF 파일 생성 중...!";
                        }
                        ProgressBar01.Stop();

                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        PH_PY900_FormItemEnabled();
                        oForm.Items.Item("YY").Specific.Value = MYY;
                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }

                    if (pVal.ItemUID == "Btn_eMail")
                    {
                        MYY = oForm.Items.Item("YY").Specific.Value.ToString().Trim(); //년월

                        ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("eMail 메일전송", 50, false);
                        oMat1.FlushToDataSource();
                        for (int i = 0; i <= oMat1.VisualRowCount - 1; i++)
                        {
                            if (oDS_PH_PY900B.GetValue("U_SendYN", i).ToString().Trim() != "Y")
                            {
                                if (!string.IsNullOrEmpty(oDS_PH_PY900B.GetValue("U_SaveYN", i).ToString().Trim()))
                                {
                                    p_MSTCOD = oDS_PH_PY900B.GetValue("U_MSTCOD", i).ToString().Trim();
                                    if (Send_EMail(p_MSTCOD) == false)//사번
                                    {
                                        errMessage = "전송 중 오류가 발생했습니다.";
                                        throw new Exception();
                                    }
                                }
                            }
                            ProgressBar01.Value += 1;
                            ProgressBar01.Text = ProgressBar01.Value + "/" + (oMat1.VisualRowCount) + "건 eMail전송중...!";
                        }
                        ProgressBar01.Stop();

                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        PH_PY900_FormItemEnabled();
                        oForm.Items.Item("YY").Specific.Value = MYY;
                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                    PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
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
                                oMat1.SelectRow(pVal.Row, true, false);
                            }
                            break;
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
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
        }

        /// <summary>
        /// Raise_EVENT_MATRIX_LOAD
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    oMat1.LoadFromDataSource();
                    PH_PY900_FormItemEnabled();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY900A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY900B);
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
        /// ROW_DELETE(Raise_FormMenuEvent에서 호출), 해당 클래스에서는 사용되지 않음
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pval"></param>
        /// <param name="BubbleEvent"></param>
        /// <param name="oMat"></param>
        /// <param name="DBData"></param>
        /// <param name="CheckField"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.IMenuEvent pval, bool BubbleEvent, SAPbouiCOM.Matrix oMat, SAPbouiCOM.DBDataSource DBData, string CheckField)
        {
            int i = 0;

            try
            {
                if (oLastColRow > 0)
                {
                    if (pval.BeforeAction == true)
                    {

                    }
                    else if (pval.BeforeAction == false)
                    {
                        if (oMat.RowCount != oMat.VisualRowCount)
                        {
                            oMat.FlushToDataSource();

                            while (i <= DBData.Size - 1)
                            {
                                if (string.IsNullOrEmpty(DBData.GetValue(CheckField, i)))
                                {
                                    DBData.RemoveRecord(i);
                                    i = 0;
                                }
                                else
                                {
                                    i += 1;
                                }
                            }

                            for (i = 0; i <= DBData.Size; i++)
                            {
                                DBData.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                            }

                            oMat.LoadFromDataSource();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ROW_DELETE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

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
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY900A", "Code"); //접속자 권한에 따른 사업장 보기
                            PH_PY900_FormItemEnabled();
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY900_FormItemEnabled();
                            PH_PY900_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY900_FormItemEnabled();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY900_FormItemEnabled();
                            PH_PY900_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY900_FormItemEnabled();
                            CLTCOD = oDS_PH_PY900A.GetValue("U_CLTCOD", 0).ToString().Trim();
                            YY = codeHelpClass.Right(oDS_PH_PY900A.GetValue("U_YY", 0).ToString().Trim(), 4);
                            break;
                        case "1293": //행삭제
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

