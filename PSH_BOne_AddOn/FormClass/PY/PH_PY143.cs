using System;
using System.IO;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 연말정산 징수 분할등록
    /// </summary>
    internal class PH_PY143 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY143A; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PH_PY143B; //등록라인

        private string oLastItemUID; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY143.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY143_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY143");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PH_PY143_CreateItems();
                PH_PY143_SetDocEntry();
                PH_PY143_FormItemEnabled();
                PH_PY143_EnableMenus();
                PH_PY143_ComboBox_Setting();
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
        private void PH_PY143_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oDS_PH_PY143A = oForm.DataSources.DBDataSources.Item("@PH_PY143A");
                oDS_PH_PY143B = oForm.DataSources.DBDataSources.Item("@PH_PY143B");

                // 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                //사업장
                sQry = "SELECT BPLId, BPLName From[OBPL] order by 1";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("CLTCOD").Specific, "N");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //해당년도
                oForm.Items.Item("YM").Specific.Value = Convert.ToString(DateTime.Now.Year - 1);

                // 구분
                oForm.DataSources.UserDataSources.Add("Gubun", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Gubun").Specific.ValidValues.Add("%", "총합");
                oForm.Items.Item("Gubun").Specific.ValidValues.Add("1", "1회차(2월)");
                oForm.Items.Item("Gubun").Specific.ValidValues.Add("2", "2회차(3월)");
                oForm.Items.Item("Gubun").Specific.ValidValues.Add("3", "3회차(4월)");
                oForm.Items.Item("Gubun").DisplayDesc = true;
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PH_PY143_SetDocEntry
        /// </summary>
        private void PH_PY143_SetDocEntry()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY143'", "");
                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.Value = "1";
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// PH_PY143_Add_MatrixRow
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        /// </summary>
        private void PH_PY143_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PH_PY143B.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PH_PY143B.Offset = oRow;
                oDS_PH_PY143B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY143_Add_MatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PH_PY143_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Combo_ValidValues_Insert("PH_PY143", "Mat01", "Cnt", "01", "01");
                dataHelpClass.Combo_ValidValues_Insert("PH_PY143", "Mat01", "Cnt", "02", "02");
                dataHelpClass.Combo_ValidValues_Insert("PH_PY143", "Mat01", "Cnt", "03", "03");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("Cnt"), "PH_PY143", "Mat01", "Cnt", false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// EnableMenus 메뉴설정
        /// </summary>
        private void PH_PY143_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.EnableMenu("1283", false);                // 삭제
                oForm.EnableMenu("1286", false);                // 닫기
                oForm.EnableMenu("1287", false);                // 복제
                oForm.EnableMenu("1285", false);                // 복원
                oForm.EnableMenu("1284", true);                // 취소
                oForm.EnableMenu("1293", false);                // 행삭제
                oForm.EnableMenu("1281", true);
                oForm.EnableMenu("1282", true);
                dataHelpClass.SetEnableMenus(oForm, false, false, false, false, false, true, true, true, true, true, false, false, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY143_MTX01
        /// </summary>
        private void PH_PY143_MTX01()
        {
            int i;
            string sQry;
            string errMessage = string.Empty;
            string Param01;
            string Param02;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                Param02 = oForm.Items.Item("YM").Specific.Value;

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                sQry = "EXEC PH_PY143_01 '" + Param01 + "','" + Param02 +"'";
                oRecordSet.DoQuery(sQry);
                
                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();
                oDS_PH_PY143B.Clear(); //추가

                if (oRecordSet.RecordCount == 0)
                {
                    errMessage = "결과값이 존재하지않습니다.";
                    oMat01.Clear();
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY143B.Size)
                    {
                        oDS_PH_PY143B.InsertRecord(i);
                    }
                    oMat01.AddRow();

                    oDS_PH_PY143B.Offset = i;
                    oDS_PH_PY143B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY143B.SetValue("U_TeamName", i, oRecordSet.Fields.Item("TeamName").Value);
                    oDS_PH_PY143B.SetValue("U_RspName", i, oRecordSet.Fields.Item("RspName").Value);
                    oDS_PH_PY143B.SetValue("U_MSTCOD", i, oRecordSet.Fields.Item("Code").Value);
                    oDS_PH_PY143B.SetValue("U_MSTNAM", i, oRecordSet.Fields.Item("U_FullName").Value);
                    oDS_PH_PY143B.SetValue("U_NStatus", i, oRecordSet.Fields.Item("NStatus").Value);
                    oDS_PH_PY143B.SetValue("U_TSAMT", i, oRecordSet.Fields.Item("TSAMT").Value);
                    oDS_PH_PY143B.SetValue("U_TJAMT", i, oRecordSet.Fields.Item("TJAMT").Value);
                    oDS_PH_PY143B.SetValue("U_TTotal", i, oRecordSet.Fields.Item("TTotal").Value);
                    oDS_PH_PY143B.SetValue("U_Cnt", i, oRecordSet.Fields.Item("Cnt").Value);
                    oDS_PH_PY143B.SetValue("U_TSAMT1", i, oRecordSet.Fields.Item("TSAMT1").Value);
                    oDS_PH_PY143B.SetValue("U_TJAMT1", i, oRecordSet.Fields.Item("TJAMT1").Value);
                    oDS_PH_PY143B.SetValue("U_TTotal1", i, oRecordSet.Fields.Item("TTotal1").Value);
                   
                    oRecordSet.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();
            }
            catch (System.Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    ProgressBar01.Stop();
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY143_MTX01:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY143_FormItemEnabled()
        {
            double STotal = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    PH_PY143_SetDocEntry();
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                    oForm.Items.Item("Gubun").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
                    for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        STotal += double.Parse(oDS_PH_PY143B.GetValue("U_TTotal", i).ToString().Trim());
                    }
                    string TotalSIL = String.Format("{0:#,###}", STotal); //자릿값변환
                    oForm.Items.Item("STotal").Specific.Value = TotalSIL;
                    TotalSIL = oForm.Items.Item("STotal").Specific.Value;
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                    oForm.Items.Item("Gubun").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
                    for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        STotal += double.Parse(oDS_PH_PY143B.GetValue("U_TTotal", i).ToString().Trim());
                    }
                    string TotalSIL = String.Format("{0:#,###}", STotal); //자릿값변환
                    oForm.Items.Item("STotal").Specific.Value = TotalSIL;
                    TotalSIL = oForm.Items.Item("STotal").Specific.Value;
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("YM").Enabled = false;
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                    oForm.Items.Item("Gubun").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
                    for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        STotal += double.Parse(oDS_PH_PY143B.GetValue("U_TTotal", i).ToString().Trim());
                    }
                    string TotalSIL = String.Format("{0:#,###}", STotal); //자릿값변환
                    oForm.Items.Item("STotal").Specific.Value = TotalSIL;
                    TotalSIL = oForm.Items.Item("STotal").Specific.Value;
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY143_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY143A_DataValidCheck()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            try
            {
               
                //년도
                if (string.IsNullOrEmpty(oForm.Items.Item("YM").Specific.Value.Trim()))
                {
                    errMessage = "년도는 필수입니다.";
                    throw new System.Exception();
                }
                returnValue = true;
            }
            catch (System.Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                    return returnValue;
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY143_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
            }
            return returnValue;
        }

        /// <summary>
        /// Raise_EVENT_COMBO_SELECT
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string errmsg = string.Empty;
            try
            {
                string sQry;
                string errMessage = string.Empty;
                double STotal = 0;
                SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    oMat01.FlushToDataSource();
                    string Para01;
                    float Para02;
                    float Para03;
                    if(pVal.ItemUID == "Gubun")
                    {
                        if (oForm.Items.Item("Gubun").Specific.Value.ToString().Trim() == "%")
                        {
                            for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                STotal += double.Parse(oDS_PH_PY143B.GetValue("U_TTotal", i).ToString().Trim());
                            }
                        }
                        else if (oForm.Items.Item("Gubun").Specific.Value.ToString().Trim() == "1")
                        {
                            for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                STotal += double.Parse(oDS_PH_PY143B.GetValue("U_TTotal1", i).ToString().Trim());
                            }
                        }
                        else if (oForm.Items.Item("Gubun").Specific.Value.ToString().Trim() == "2")
                        {
                            for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                STotal += double.Parse(oDS_PH_PY143B.GetValue("U_TTotal2", i).ToString().Trim());
                            }
                        }
                        else if (oForm.Items.Item("Gubun").Specific.Value.ToString().Trim() == "3")
                        {
                            for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                STotal += double.Parse(oDS_PH_PY143B.GetValue("U_TTotal3", i).ToString().Trim());
                            }
                        }
                        string TotalSIL = String.Format("{0:#,###}", STotal); //자릿값변환
                        oForm.Items.Item("STotal").Specific.Value = TotalSIL;
                        TotalSIL = oForm.Items.Item("STotal").Specific.Value;
                    }
                    if (pVal.ItemUID == "Mat01" && pVal.ColUID == "Cnt")
                    {
                        if (oMat01.Columns.Item("Cnt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() == "01")
                        {
                            Para01 = "01";
                            Para02 = float.Parse(oMat01.Columns.Item("TSAMT").Cells.Item(pVal.Row).Specific.Value);
                            Para03 = float.Parse(oMat01.Columns.Item("TJAMT").Cells.Item(pVal.Row).Specific.Value);
                            sQry = "EXEC PH_PY143_02 '" + Para01 + "','" + Para02 + "','" + Para03 + "'";
                            oRecordSet.DoQuery(sQry);
                        }
                        else if (oMat01.Columns.Item("Cnt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() == "02")
                        {
                            Para01 = "02";
                            Para02 = float.Parse(oMat01.Columns.Item("TSAMT").Cells.Item(pVal.Row).Specific.Value);
                            Para03 = float.Parse(oMat01.Columns.Item("TJAMT").Cells.Item(pVal.Row).Specific.Value);
                            sQry = "EXEC PH_PY143_02 '" + Para01 + "','" + Para02 + "','" + Para03 + "'";
                            oRecordSet.DoQuery(sQry);
                        }
                        else if (oMat01.Columns.Item("Cnt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() == "03")
                        {
                            Para01 = "03";
                            Para02 = float.Parse(oMat01.Columns.Item("TSAMT").Cells.Item(pVal.Row).Specific.Value);
                            Para03 = float.Parse(oMat01.Columns.Item("TJAMT").Cells.Item(pVal.Row).Specific.Value);
                            sQry = "EXEC PH_PY143_02 '" + Para01 + "','" + Para02 + "','" + Para03 + "'";
                            oRecordSet.DoQuery(sQry);
                        }
                        oDS_PH_PY143B.SetValue("U_TSAMT1", pVal.Row - 1, oRecordSet.Fields.Item("TSAMT1").Value);
                        oDS_PH_PY143B.SetValue("U_TJAMT1", pVal.Row -1, oRecordSet.Fields.Item("TJAMT1").Value);
                        oDS_PH_PY143B.SetValue("U_TTotal1", pVal.Row - 1, oRecordSet.Fields.Item("TTotal1").Value);
                        oDS_PH_PY143B.SetValue("U_TSAMT2", pVal.Row - 1, oRecordSet.Fields.Item("TSAMT2").Value);
                        oDS_PH_PY143B.SetValue("U_TJAMT2", pVal.Row - 1, oRecordSet.Fields.Item("TJAMT2").Value);
                        oDS_PH_PY143B.SetValue("U_TTotal2", pVal.Row - 1, oRecordSet.Fields.Item("TTotal2").Value);
                        oDS_PH_PY143B.SetValue("U_TSAMT3", pVal.Row - 1, oRecordSet.Fields.Item("TSAMT3").Value);
                        oDS_PH_PY143B.SetValue("U_TJAMT3", pVal.Row - 1, oRecordSet.Fields.Item("TJAMT3").Value);
                        oDS_PH_PY143B.SetValue("U_TTotal3", pVal.Row - 1, oRecordSet.Fields.Item("TTotal3").Value);
                    }
                    oMat01.LoadFromDataSource();
                    oMat01.AutoResizeColumns();
                    oForm.Update();
                }
            }
            catch (Exception ex)
            {
                if (errmsg != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errmsg);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY126_DataApply
        /// </summary>
        /// <param name="Gubun"></param>
        /// <param name="MSTCOD"></param>
        /// <param name="AMT"></param>
        /// <returns></returns>
        private bool PH_PY143_DataApply(string Gubun, string MSTCOD, string AMT)
        {
            bool returnValue = false;
            string sQry;
            string YY;
            string AMTCode;
            string CLTCOD;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                oMat01.FlushToDataSource();
                YY = codeHelpClass.Right((Int64.Parse(oForm.Items.Item("YM").Specific.Value.ToString().Trim()) + 1).ToString(),2);
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                //소득세정산코드 : G06
                if (Gubun == "1")
                {
                    //2월 급상여변동자료에 등록
                    sQry = "select U_Sequence from [@PH_PY109Z] where U_PDCode ='G06' AND code ='" + CLTCOD + YY + "02111'";
                    oRecordSet.DoQuery(sQry);
                    if (oRecordSet.RecordCount == 0)
                    {
                        errMessage = "급상여변동자료가 등록되지 않았습니다. 등록을 진행하세요.";
                        throw new Exception();
                    }
                    AMTCode = Convert.ToString(Convert.ToDouble("0") + oRecordSet.Fields.Item(0).Value.ToString().Trim());

                    sQry = " UPDATE [@PH_PY109B] SET U_AMT" + AMTCode + "= ISNULL(U_AMT" + AMTCode + ",0) + " + AMT;
                    sQry += "  from [@PH_PY109B] WHERE U_MSTCOD = '"+ MSTCOD + "' AND " + "Code ='" + CLTCOD + YY + "02111'";
                    oRecordSet.DoQuery(sQry);

                    sQry = " update [@PH_PY143B] set U_YN1 = 'Y' where DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND U_MSTCOD = '" + MSTCOD + "'";
                    oRecordSet.DoQuery(sQry);
                }
                else if (Gubun == "2")
                {
                    //3월 급상여변동자료에 등록
                    sQry = "select U_Sequence from [@PH_PY109Z] where U_PDCode ='G06' AND code ='" + CLTCOD + YY + "03111'";
                    oRecordSet.DoQuery(sQry);
                    if (oRecordSet.RecordCount == 0)
                    {
                        errMessage = "급상여변동자료가 등록되지 않았습니다. 등록을 진행하세요.";
                        throw new Exception();
                    }
                    AMTCode = Convert.ToString(Convert.ToDouble("0") + oRecordSet.Fields.Item(0).Value.ToString().Trim());

                    sQry = " UPDATE [@PH_PY109B] SET U_AMT" + AMTCode + "= ISNULL(U_AMT" + AMTCode + ",0) + " + AMT;
                    sQry += "  from [@PH_PY109B] WHERE U_MSTCOD = '" + MSTCOD + "' AND " + "Code ='" + CLTCOD + YY + "03111'";
                    oRecordSet.DoQuery(sQry);

                    sQry = " update [@PH_PY143B] set U_YN2 = 'Y' where DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND U_MSTCOD = '" + MSTCOD + "'";
                    oRecordSet.DoQuery(sQry);
                }
                else if (Gubun == "3")
                {
                    //4월 급상여변동자료에 등록
                    sQry = "select U_Sequence from [@PH_PY109Z] where U_PDCode ='G06' AND code ='" + CLTCOD + YY + "04111'";
                    oRecordSet.DoQuery(sQry);
                    if (oRecordSet.RecordCount == 0)
                    {
                        errMessage = "급상여변동자료가 등록되지 않았습니다. 등록을 진행하세요.";
                        throw new Exception();
                    }
                    AMTCode = Convert.ToString(Convert.ToDouble("0") + oRecordSet.Fields.Item(0).Value.ToString().Trim());

                    sQry = " UPDATE [@PH_PY109B] SET U_AMT" + AMTCode + "= ISNULL(U_AMT" + AMTCode + ",0) + " + AMT;
                    sQry += "  from [@PH_PY109B] WHERE  U_MSTCOD = '" + MSTCOD + "' AND " + "Code = '" + CLTCOD + YY + "04111'";
                    oRecordSet.DoQuery(sQry);

                    sQry = " update [@PH_PY143B] set U_YN3 = 'Y' where DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND U_MSTCOD = '" + MSTCOD + "'";
                    oRecordSet.DoQuery(sQry);
                }
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
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
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
        /// <param name="Gubun"></param>
        /// <param name="MSTCOD"></param>
        /// <param name="AMT"></param>
        /// <returns></returns>
        private bool PH_PY143_DataCancel(string Gubun, string MSTCOD, string AMT)
        {
            
            bool returnValue = false;
            string sQry;
            string YY;
            string CLTCOD;
            string AMTCode;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                oMat01.FlushToDataSource();

                oMat01.FlushToDataSource();
                YY = codeHelpClass.Right((Int64.Parse(oForm.Items.Item("YM").Specific.Value.ToString().Trim()) + 1).ToString(), 2);
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                //소득세정산코드 : G06
                if (Gubun == "1")
                {
                    //2월 급상여변동자료에 등록
                    sQry = "select U_Sequence from [@PH_PY109Z] where U_PDCode ='G06' AND code ='" + CLTCOD + YY + "02111'";
                    oRecordSet.DoQuery(sQry);
                    if (oRecordSet.RecordCount == 0)
                    {
                        errMessage = "급상여변동자료가 등록되지 않았습니다. 등록을 진행하세요.";
                        throw new Exception();
                    }
                    AMTCode = Convert.ToString(Convert.ToDouble("0") + oRecordSet.Fields.Item(0).Value.ToString().Trim());

                    sQry = " UPDATE [@PH_PY109B] SET U_AMT" + AMTCode + "= ISNULL(U_AMT" + AMTCode + ",0) - " + AMT;
                    sQry += "  from [@PH_PY109B] WHERE U_MSTCOD = '" + MSTCOD + "' AND " + "Code ='" + CLTCOD + YY + "02111'";
                    oRecordSet.DoQuery(sQry);

                    sQry = " update [@PH_PY143B] set U_YN1 = NULL where DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND U_MSTCOD = '" + MSTCOD + "'";
                    oRecordSet.DoQuery(sQry);
                }
                else if (Gubun == "2")
                {
                    //3월 급상여변동자료에 등록
                    sQry = "select U_Sequence from [@PH_PY109Z] where U_PDCode ='G06' AND code ='" + CLTCOD + YY + "03111'";
                    oRecordSet.DoQuery(sQry);
                    if (oRecordSet.RecordCount == 0)
                    {
                        errMessage = "급상여변동자료가 등록되지 않았습니다. 등록을 진행하세요.";
                        throw new Exception();
                    }
                    AMTCode = Convert.ToString(Convert.ToDouble("0") + oRecordSet.Fields.Item(0).Value.ToString().Trim());

                    sQry = " UPDATE [@PH_PY109B] SET U_AMT" + AMTCode + "= ISNULL(U_AMT" + AMTCode + ",0) - " + AMT;
                    sQry += "  from [@PH_PY109B] WHERE U_MSTCOD = '" + MSTCOD + "' AND " + "Code ='" + CLTCOD + YY + "03111'";
                    oRecordSet.DoQuery(sQry);

                    sQry = " update [@PH_PY143B] set U_YN2 = NULL where DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND U_MSTCOD = '" + MSTCOD + "'";
                    oRecordSet.DoQuery(sQry);
                }
                else if (Gubun == "3")
                {
                    //4월 급상여변동자료에 등록
                    sQry = "select U_Sequence from [@PH_PY109Z] where U_PDCode ='G06' AND code ='" + CLTCOD + YY + "04111'";
                    oRecordSet.DoQuery(sQry);
                    if (oRecordSet.RecordCount == 0)
                    {
                        errMessage = "급상여변동자료가 등록되지 않았습니다. 등록을 진행하세요.";
                        throw new Exception();
                    }
                    AMTCode = Convert.ToString(Convert.ToDouble("0") + oRecordSet.Fields.Item(0).Value.ToString().Trim());

                    sQry = " UPDATE [@PH_PY109B] SET U_AMT" + AMTCode + "= ISNULL(U_AMT" + AMTCode + ",0) - " + AMT;
                    sQry += "  from [@PH_PY109B] WHERE  U_MSTCOD = '" + MSTCOD + "' AND " + "Code = '" + CLTCOD + YY + "04111'";
                    oRecordSet.DoQuery(sQry);

                    sQry = " update [@PH_PY143B] set U_YN3 = NULL where DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND U_MSTCOD = '" + MSTCOD + "'";
                    oRecordSet.DoQuery(sQry);
                }
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
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
            return returnValue;
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

                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                //    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                if (pVal.ItemUID == "Mat01")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID = pVal.ItemUID;
                        oLastColUID = pVal.ColUID;
                        oLastColRow = pVal.Row;
                    }
                }
                else
                {
                    oLastItemUID = pVal.ItemUID;
                    oLastColUID = "";
                    oLastColRow = 0;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                     PH_PY143_FormItemEnabled();
                }
            }
            catch (System.Exception ex)
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY143A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY143B);
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            string MSTCOD;
            string AMT;
            string sVersion;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgressBar01 = null; 
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    sVersion = oForm.Items.Item("DocEntry").Specific.Value;

                    //조회
                    if (pVal.ItemUID == "Btn01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY143A_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            //동일한거있는지 확인
                            sQry = "SELECT COUNT(*) FROM [@PH_PY143A] WHERE Canceled <> 'Y' AND U_CLTCOD ='" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'AND U_YM ='" + oForm.Items.Item("YM").Specific.Value.Trim() + "'";
                            oRecordSet.DoQuery(sQry);
                            if (oRecordSet.Fields.Item(0).Value != 0)
                            {
                                errMessage = "동일한 문서가 있습니다. 확인하세요.";
                                throw new System.Exception();
                            }
                            PH_PY143_MTX01();
                        }
                    }
                    //급여에 적용
                    if (pVal.ItemUID == "Btn_UPDATE")
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("급상여 변동자료에 적용(등록) 하시겠습니까?.", 2, "Yes", "No") == 1)
                        {
                            
                            if (oForm.Items.Item("Gubun").Specific.Value.ToString().Trim() == "%")
                            {
                                errMessage = "회차 구분을 선택하세요";
                                BubbleEvent = false;
                                throw new System.Exception();
                            }
                            else if (oForm.Items.Item("Gubun").Specific.Value.ToString().Trim() == "1")
                            {
                                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    if (oDS_PH_PY143B.GetValue("U_YN1", i).ToString().Trim() == "Y" || oDS_PH_PY143B.GetValue("U_TSAMT1", i).ToString().Trim() == "0.0")
                                    {
                                        //이미 적용된 것들은 넘어감.
                                    }
                                    else
                                    {
                                        MSTCOD = oDS_PH_PY143B.GetValue("U_MSTCOD", i).ToString().Trim();
                                        AMT = oDS_PH_PY143B.GetValue("U_TTotal1", i).ToString().Trim();
                                        if (PH_PY143_DataApply("1", MSTCOD, AMT) == false)
                                        {
                                            BubbleEvent = false;
                                            return;
                                        }

                                    }
                                }
                            }
                            else if (oForm.Items.Item("Gubun").Specific.Value.ToString().Trim() == "2")
                            {
                                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    if (oDS_PH_PY143B.GetValue("U_YN2", i).ToString().Trim() == "Y" || oDS_PH_PY143B.GetValue("U_TSAMT2", i).ToString().Trim() == "0.0")
                                    {
                                        //이미 적용된 것들은 넘어감.
                                    }
                                    else
                                    {
                                        MSTCOD = oDS_PH_PY143B.GetValue("U_MSTCOD", i).ToString().Trim();
                                        AMT = oDS_PH_PY143B.GetValue("U_TTotal2", i).ToString().Trim();
                                        if (PH_PY143_DataApply("2", MSTCOD, AMT) == false)
                                        {
                                            BubbleEvent = false;
                                            return;
                                        }

                                    }
                                }
                            }
                            else if (oForm.Items.Item("Gubun").Specific.Value.ToString().Trim() == "3")
                            {
                                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    if (oDS_PH_PY143B.GetValue("U_YN3", i).ToString().Trim() == "Y" || oDS_PH_PY143B.GetValue("U_TSAMT3", i).ToString().Trim() == "0.0")
                                    {
                                        //이미 적용된 것들은 넘어감.
                                    }
                                    else
                                    {
                                        MSTCOD = oDS_PH_PY143B.GetValue("U_MSTCOD", i).ToString().Trim();
                                        AMT = oDS_PH_PY143B.GetValue("U_TTotal3", i).ToString().Trim();
                                        if (PH_PY143_DataApply("3", MSTCOD, AMT) == false)
                                        {
                                            BubbleEvent = false;
                                            return;
                                        }

                                    }
                                }
                            }
                            PSH_Globals.SBO_Application.StatusBar.SetText("급상여변동 자료에 금액이 적용(등록) 되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        PH_PY143_FormItemEnabled();
                        oForm.Items.Item("DocEntry").Specific.Value = sVersion;
                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    //급여에 적용취소
                    if (pVal.ItemUID == "Btn_CANCEL")
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("급상여 변동자료에 적용(취소) 하시겠습니까?.", 2, "Yes", "No") == 1)
                        {
                            if (oForm.Items.Item("Gubun").Specific.Value.ToString().Trim() == "%")
                            {
                                errMessage = "회차 구분을 선택하세요";
                                BubbleEvent = false;
                                throw new System.Exception();
                            }
                            else if (oForm.Items.Item("Gubun").Specific.Value.ToString().Trim() == "1")
                            {
                                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    if (oDS_PH_PY143B.GetValue("U_YN1", i).ToString().Trim() == "Y")
                                    {
                                        MSTCOD = oDS_PH_PY143B.GetValue("U_MSTCOD", i).ToString().Trim();
                                        AMT = oDS_PH_PY143B.GetValue("U_TTotal1", i).ToString().Trim();
                                        if (PH_PY143_DataCancel("1", MSTCOD, AMT) == false)
                                        {
                                            BubbleEvent = false;
                                            return;
                                        }
                                    }
                                }
                            }
                            else if (oForm.Items.Item("Gubun").Specific.Value.ToString().Trim() == "2")
                            {
                                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    if (oDS_PH_PY143B.GetValue("U_YN2", i).ToString().Trim() == "Y")
                                    {
                                        MSTCOD = oDS_PH_PY143B.GetValue("U_MSTCOD", i).ToString().Trim();
                                        AMT = oDS_PH_PY143B.GetValue("U_TTotal2", i).ToString().Trim();
                                        if (PH_PY143_DataCancel("2", MSTCOD, AMT) == false)
                                        {
                                            BubbleEvent = false;
                                            return;
                                        }
                                    }
                                }
                            }
                            else if (oForm.Items.Item("Gubun").Specific.Value.ToString().Trim() == "3")
                            {
                                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    if (oDS_PH_PY143B.GetValue("U_YN3", i).ToString().Trim() == "Y")
                                    {
                                        MSTCOD = oDS_PH_PY143B.GetValue("U_MSTCOD", i).ToString().Trim();
                                        AMT = oDS_PH_PY143B.GetValue("U_TTotal3", i).ToString().Trim();
                                        if (PH_PY143_DataCancel("3", MSTCOD, AMT) == false)
                                        {
                                            BubbleEvent = false;
                                            return;
                                        }
                                    }
                                }
                            }
                            PSH_Globals.SBO_Application.StatusBar.SetText("급상여변동 자료에 금액이 적용(취소) 되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        PH_PY143_FormItemEnabled();
                        oForm.Items.Item("DocEntry").Specific.Value = sVersion;
                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    //추가
                    if (pVal.ItemUID == "1") 
                    { 
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY143A_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                            if (oDS_PH_PY143B.Size < 1)
                            {
                                errMessage = "조회 누르르고 추가하세오!";
                                BubbleEvent = false;
                                throw new System.Exception();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        {
                        }
                    }

                    else if (pVal.BeforeAction == false)
                    {
                        if (pVal.ItemUID == "1")
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY143_FormItemEnabled();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
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
                oForm.Freeze(false);
                if (ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// FormMenuEvent
        /// <summary>
        /// 메뉴이벤트
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
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
                                }

                                oMat01.FlushToDataSource();
                                oDS_PH_PY143B.RemoveRecord(oDS_PH_PY143B.Size - 1);
                                oMat01.Clear();
                                oMat01.LoadFromDataSource();

                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("LineNum").Cells.Item(oMat01.RowCount).Specific.Value))
                                {
                                    PH_PY143_AddMatrixRow(oMat01.RowCount, false);
                                }
                            }
                            break;
                        case "1281": //찾기
                            PH_PY143_FormItemEnabled();
                            break;
                        case "1282": //추가
                            PH_PY143_FormItemEnabled();
                            PH_PY143_SetDocEntry();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PH_PY143_FormItemEnabled();
                            break;
                        case "1287": //복제
                            break;
                    }
                }
            }
            catch (System.Exception ex)
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

