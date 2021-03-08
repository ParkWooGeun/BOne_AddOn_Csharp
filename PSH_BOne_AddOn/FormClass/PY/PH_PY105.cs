using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 호봉표등록
    /// </summary>
    internal class PH_PY105 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY105A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY105B;
        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;
        private bool CheckDataApply; //적용버튼 실행여부
        private string CLTCOD; //사업장
        private string YM; //적용연월

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY105.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }
                oFormUniqueID = "PH_PY105_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY105");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                PH_PY105_CreateItems();
                PH_PY105_EnableMenus();
                PH_PY105_SetDocument(oFormDocEntry01);
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
        /// <returns></returns>
        private void PH_PY105_CreateItems()
        {
            int i;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                oDS_PH_PY105A = oForm.DataSources.DBDataSources.Item("@PH_PY105A"); //헤더
                oDS_PH_PY105B = oForm.DataSources.DBDataSources.Item("@PH_PY105B"); //라인

                oMat1 = oForm.Items.Item("Mat1").Specific;
                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                CheckDataApply = false;
                
                oForm.Items.Item("CLTCOD").DisplayDesc = true; //사업장
                
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P129' AND U_UseYN= 'Y'"; //직급
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oMat1.Columns.Item("JIGCOD").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                        oRecordSet.MoveNext();
                    }
                }
                oMat1.Columns.Item("JIGCOD").DisplayDesc = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY105_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true);  //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true);  //행삭제
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
        /// PH_PY105_SetDocument
        /// </summary>
        private void PH_PY105_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY105_FormItemEnabled();
                    PH_PY105_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY105_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY105_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = false;

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.Items.Item("YM").Specific.Value = DateTime.Now.ToString("yyyyMM"); //귀속년월

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = false;

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("YM").Enabled = false;
                    oForm.Items.Item("Comments").Enabled = false;

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
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
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY105_AddMatrixRow()
        {
            try
            {
                oForm.Freeze(true);

                oMat1.FlushToDataSource();
                int oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY105B.GetValue("U_HOBCOD", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY105B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY105B.InsertRecord(oRow);
                        }
                        oDS_PH_PY105B.Offset = oRow;
                        oDS_PH_PY105B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY105B.SetValue("U_JIGCOD", oRow, "");
                        oDS_PH_PY105B.SetValue("U_HOBCOD", oRow, "");
                        oDS_PH_PY105B.SetValue("U_HOBNAM", oRow, "");
                        oDS_PH_PY105B.SetValue("U_STDAMT", oRow, "0");
                        oDS_PH_PY105B.SetValue("U_BNSAMT", oRow, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT01", oRow, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT02", oRow, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT03", oRow, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT04", oRow, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT05", oRow, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT06", oRow, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT07", oRow, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT08", oRow, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT09", oRow, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT10", oRow, "0");
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY105B.Offset = oRow - 1;
                        oDS_PH_PY105B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY105B.SetValue("U_JIGCOD", oRow - 1, "");
                        oDS_PH_PY105B.SetValue("U_HOBCOD", oRow - 1, "");
                        oDS_PH_PY105B.SetValue("U_HOBNAM", oRow - 1, "");
                        oDS_PH_PY105B.SetValue("U_STDAMT", oRow - 1, "0");
                        oDS_PH_PY105B.SetValue("U_BNSAMT", oRow - 1, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT01", oRow - 1, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT02", oRow - 1, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT03", oRow - 1, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT04", oRow - 1, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT05", oRow - 1, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT06", oRow - 1, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT07", oRow - 1, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT08", oRow - 1, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT09", oRow - 1, "0");
                        oDS_PH_PY105B.SetValue("U_EXTAMT10", oRow - 1, "0");
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY105B.Offset = oRow;
                    oDS_PH_PY105B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY105B.SetValue("U_JIGCOD", oRow, "");
                    oDS_PH_PY105B.SetValue("U_HOBCOD", oRow, "");
                    oDS_PH_PY105B.SetValue("U_HOBNAM", oRow, "");
                    oDS_PH_PY105B.SetValue("U_STDAMT", oRow, "0");
                    oDS_PH_PY105B.SetValue("U_BNSAMT", oRow, "0");
                    oDS_PH_PY105B.SetValue("U_EXTAMT01", oRow, "0");
                    oDS_PH_PY105B.SetValue("U_EXTAMT02", oRow, "0");
                    oDS_PH_PY105B.SetValue("U_EXTAMT03", oRow, "0");
                    oDS_PH_PY105B.SetValue("U_EXTAMT04", oRow, "0");
                    oDS_PH_PY105B.SetValue("U_EXTAMT05", oRow, "0");
                    oDS_PH_PY105B.SetValue("U_EXTAMT06", oRow, "0");
                    oDS_PH_PY105B.SetValue("U_EXTAMT07", oRow, "0");
                    oDS_PH_PY105B.SetValue("U_EXTAMT08", oRow, "0");
                    oDS_PH_PY105B.SetValue("U_EXTAMT09", oRow, "0");
                    oDS_PH_PY105B.SetValue("U_EXTAMT10", oRow, "0");
                    oMat1.LoadFromDataSource();
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
        /// 엑셀자료 업로드
        /// </summary>
        [STAThread]
        private void PH_PY105_Excel_Upload()
        {
            int rowCount;
            int loopCount;
            string sFile;
            bool sucessFlag = false;
            short columnCount = 15; //엑셀 컬럼수

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

            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("시작!", xlRow.Count, false);

            oForm.Freeze(true);

            oMat1.Clear();
            oMat1.FlushToDataSource();
            oMat1.LoadFromDataSource();

            try
            {
                for (rowCount = 2; rowCount <= xlRow.Count; rowCount++)
                {
                    if (rowCount - 2 != 0)
                    {
                        oDS_PH_PY105B.InsertRecord(rowCount - 2);
                    }

                    Microsoft.Office.Interop.Excel.Range[] r = new Microsoft.Office.Interop.Excel.Range[columnCount + 1];

                    for (loopCount = 1; loopCount <= columnCount; loopCount++)
                    {
                        r[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[rowCount, loopCount];
                    }

                    oDS_PH_PY105B.Offset = rowCount - 2;
                    oDS_PH_PY105B.SetValue("U_LineNum", rowCount - 2, Convert.ToString(rowCount - 1));
                    oDS_PH_PY105B.SetValue("U_JIGCOD", rowCount - 2, Convert.ToString(r[1].Value)); //직급
                    oDS_PH_PY105B.SetValue("U_HOBCOD", rowCount - 2, Convert.ToString(r[2].Value)); //호봉코드
                    oDS_PH_PY105B.SetValue("U_HOBNAM", rowCount - 2, Convert.ToString(r[3].Value)); //호봉명
                    oDS_PH_PY105B.SetValue("U_STDAMT", rowCount - 2, Convert.ToString(r[4].Value)); //급여기본
                    oDS_PH_PY105B.SetValue("U_BNSAMT", rowCount - 2, Convert.ToString(r[5].Value)); //상여기본
                    oDS_PH_PY105B.SetValue("U_EXTAMT01", rowCount - 2, Convert.ToString(r[6].Value)); //제수당01
                    oDS_PH_PY105B.SetValue("U_EXTAMT02", rowCount - 2, Convert.ToString(r[7].Value)); //제수당02
                    oDS_PH_PY105B.SetValue("U_EXTAMT03", rowCount - 2, Convert.ToString(r[8].Value)); //제수당03
                    oDS_PH_PY105B.SetValue("U_EXTAMT04", rowCount - 2, Convert.ToString(r[9].Value)); //제수당04
                    oDS_PH_PY105B.SetValue("U_EXTAMT05", rowCount - 2, Convert.ToString(r[10].Value)); //제수당05
                    oDS_PH_PY105B.SetValue("U_EXTAMT06", rowCount - 2, Convert.ToString(r[11].Value)); //제수당06
                    oDS_PH_PY105B.SetValue("U_EXTAMT07", rowCount - 2, Convert.ToString(r[12].Value)); //제수당07
                    oDS_PH_PY105B.SetValue("U_EXTAMT08", rowCount - 2, Convert.ToString(r[13].Value)); //제수당08
                    oDS_PH_PY105B.SetValue("U_EXTAMT09", rowCount - 2, Convert.ToString(r[14].Value)); //제수당09
                    oDS_PH_PY105B.SetValue("U_EXTAMT10", rowCount - 2, Convert.ToString(r[15].Value)); //제수당10

                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + (xlRow.Count - 1) + "건 Loding...!";

                    for (loopCount = 1; loopCount <= columnCount; loopCount++)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(r[loopCount]); //메모리 해제
                    }
                }

                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();
                oForm.Update();

                PH_PY105_AddMatrixRow();
                sucessFlag = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox("[PH_PY105_Excel_Upload_Error]" + (char)13 + ex.Message);
            }
            finally
            {
                //액셀개체 닫음
                xlapp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRow);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCell);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsh);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlshs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwbs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp);

                ProgressBar01.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);

                if (sucessFlag == true)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("엑셀 Loding 완료", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }

                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY105_DataApply
        /// </summary>
        /// <returns></returns>
        private bool PH_PY105_DataApply(string CLTCOD, string YM)
        {
            bool functionReturnValue = false;
            string sQry;
            string Tablename;
            string sTablename;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat1.FlushToDataSource();
                sTablename = "@PH_PY001A_" + YM;
                //조회용
                sQry = " SELECT Count(*) FROM SYSOBJECTS WHERE Name = '" + sTablename + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value == 0)
                {
                    Tablename = "[@PH_PY001A_" + YM + "]";
                    //테이블 생성용
                    sQry = "Exec PH_PY105_99 '" + Tablename + "'";
                    oRecordSet.DoQuery(sQry);
                }

                sQry = "";
                sQry += " Update T2 ";
                sQry += " SET T2.U_STDAMT = T1.U_STDAMT, T2.U_BNSAMT = T1.U_BNSAMT, T2.U_HOBYMM = T0.U_YM";
                sQry += " FROM [@PH_PY105A] T0";
                sQry += " INNER JOIN [@PH_PY105B] T1 ON T0.Code = T1.Code";
                sQry += " INNER JOIN [@PH_PY001A] T2 ON T2.U_JIGCOD = T1.U_JIGCOD AND T2.U_HOBONG = T1.U_HOBCOD";
                sQry += " WHERE T0.U_YM = '" + YM + "'";
                sQry += " And T2.U_status <> '5' ";
                sQry += " And Not Exists (Select * From [@PH_PY001A] T3 ";
                sQry += " Where T2.Code = T3.Code";
                sQry += " And T3.U_status <> '5'";
                sQry += " And dbo.PH_PY_PAYPEAK_YEAR(T3.U_CLTCOD,'" + YM + "',T3.Code) > 0 )";
                //호봉등록 년월에 임금피크제대상은 임금조정을 안함.
                //sQry = sQry & " AND T0.U_YM = '" & YM & "'"
                oRecordSet.DoQuery(sQry);

                PSH_Globals.SBO_Application.StatusBar.SetText("해당 직급에 대헤 인사마스터에 금액이 적용 되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY105_DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY105_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //헤더----------
                //사업장
                if (string.IsNullOrEmpty(oDS_PH_PY105A.GetValue("U_CLTCOD", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                //적용시작월
                if (string.IsNullOrEmpty(oDS_PH_PY105A.GetValue("U_YM", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("적용시작월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                //Code & Name 생성
                oDS_PH_PY105A.SetValue("Code", 0, oDS_PH_PY105A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY105A.GetValue("U_YM", 0).ToString().Trim());
                oDS_PH_PY105A.SetValue("NAME", 0, oDS_PH_PY105A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY105A.GetValue("U_YM", 0).ToString().Trim());

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (!string.IsNullOrEmpty(dataHelpClass.Get_ReData("Code", "Code", "[@PH_PY105A]", "'" + oDS_PH_PY105A.GetValue("Code", 0) + "'", "")))
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("이미 존재하는 코드입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        return functionReturnValue;
                    }
                }

                //라인----------
                if (oMat1.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                    {
                        //호봉코드
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("HOBCOD").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("호봉코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("HOBCOD").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                        //호봉명
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("HOBNAM").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("내역 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("HOBNAM").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return functionReturnValue;
                }

                oMat1.FlushToDataSource();
                if (oDS_PH_PY105B.Size > 1)
                {
                    oDS_PH_PY105B.RemoveRecord(oDS_PH_PY105B.Size - 1);
                }
                oMat1.LoadFromDataSource();

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
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
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
            string FullName = string.Empty;
            string FindYN = string.Empty;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY105_DataValidCheck() == false)
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
                                PH_PY105_DataApply(CLTCOD, YM);
                                CheckDataApply = false;
                            }
                            PH_PY105_FormItemEnabled();
                            PH_PY105_AddMatrixRow();
                        }
                    }
                    if (pVal.ItemUID == "Btn_UPLOAD")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY105_Excel_Upload);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                    if (pVal.ItemUID == "Btn_Apply")
                    {
                        CLTCOD = oDS_PH_PY105A.GetValue("U_CLTCOD", 0).ToString().Trim();
                        YM = oDS_PH_PY105A.GetValue("U_YM", 0).ToString().Trim();
                        if (oMat1.RowCount > 1)
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                CheckDataApply = true;
                                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                PH_PY105_DataApply(CLTCOD, YM);
                            }
                        }
                        else
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("호봉표 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        /// Raise_EVENT_GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    oMat1.LoadFromDataSource();

                    PH_PY105_FormItemEnabled();
                    PH_PY105_AddMatrixRow();
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
        /// Raise_EVENT_KEY_DOWN
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
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
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
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
                        if (pVal.ItemUID == "Mat1" && pVal.ColUID == "HOBCOD")
                        {
                            PH_PY105_AddMatrixRow();
                            oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                oForm.Freeze(false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY105A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY105B);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
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
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY105A", "Code"); //접속자 권한에 따른 사업장 보기
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY105_FormItemEnabled();
                            PH_PY105_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY105_FormItemEnabled();
                            PH_PY105_AddMatrixRow();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY105_FormItemEnabled();
                            PH_PY105_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY105_FormItemEnabled();
                            break;
                        case "1293": //행삭제
                            if (oMat1.RowCount != oMat1.VisualRowCount)
                            {
                                oMat1.FlushToDataSource();

                                while (i <= oDS_PH_PY105B.Size - 1)
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY105B.GetValue("U_LineNum", i)))
                                    {
                                        oDS_PH_PY105B.RemoveRecord(i);
                                        i = 0;
                                    }
                                    else
                                    {
                                        i += 1;
                                    }
                                }
                                for (i = 0; i <= oDS_PH_PY105B.Size; i++)
                                {
                                    oDS_PH_PY105B.SetValue("U_LineNum", i, (i + 1).ToString());
                                }
                                oMat1.LoadFromDataSource();
                            }
                            PH_PY105_AddMatrixRow();
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
    }
}