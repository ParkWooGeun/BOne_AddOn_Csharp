using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// Î≤†ÎÑ§?ºÏïÑ Í∏àÏï°?±Î°ù
    /// </summary>
    internal class PH_PY124 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY124A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY124B;
        private string oLastItemUID;     //?¥Îûò?§Ïóê?? ?†ÌÉù?? ÎßàÏ?Îß? ?ÑÏù¥?? UidÍ∞?
        private string oLastColUID;      //ÎßàÏ?ÎßâÏïÑ?¥ÌÖú?? Î©îÌä∏Î¶?ä§?ºÍ≤Ω?∞Ïóê ÎßàÏ?Îß? ?†ÌÉù?? Col?? UidÍ∞?
        private int oLastColRow;         //ÎßàÏ?ÎßâÏïÑ?¥ÌÖú?? Î©îÌä∏Î¶?ä§?ºÍ≤Ω?∞Ïóê ÎßàÏ?Îß? ?†ÌÉù?? RowÍ∞?
        private bool CheckDataApply; //?ÅÏö©Î≤ÑÌäº ?§Ìñâ?¨Î?
        private string CLTCOD; //?¨ÏóÖ??
        private string YM; //?ÅÏö©?∞Ïõî

        /// <summary>
        /// Form ?∏Ï∂ú
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY124.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY124_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY124");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";
                
                oForm.Freeze(true);
                PH_PY124_CreateItems();
                PH_PY124_EnableMenus();
                PH_PY124_SetDocument(oFormDocEntry01);
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //Î©îÎ™®Î¶? ?¥Ï†ú
            }
        }

        /// <summary>
        /// ?îÎ©¥ Item ?ùÏÑ±
        /// </summary>
        /// <returns></returns>
        private void PH_PY124_CreateItems()
        {
            try
            {
                oForm.Freeze(true);
                oDS_PH_PY124A = oForm.DataSources.DBDataSources.Item("@PH_PY124A");
                oDS_PH_PY124B = oForm.DataSources.DBDataSources.Item("@PH_PY124B");

                oMat1 = oForm.Items.Item("Mat1").Specific;
                
                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                CheckDataApply = false;

                oForm.Items.Item("CLTCOD").DisplayDesc = true; //?¨ÏóÖ??
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY124_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Î©îÎâ¥ ?ÑÏù¥ÏΩ? Enable
        /// </summary>
        private void PH_PY124_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true);                ////?úÍ±∞
                oForm.EnableMenu("1284", false);                ////Ï∑®ÏÜå
                oForm.EnableMenu("1293", true);                ////?âÏÇ≠??
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY124_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// ?îÎ©¥?∏ÌåÖ
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY124_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY124_FormItemEnabled();
                    PH_PY124_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY124_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY124_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ?îÎ©¥?? ?ÑÏù¥?? Enable ?§Ï†ï
        /// </summary>
        private void PH_PY124_FormItemEnabled()
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
                    oForm.Items.Item("FieldCo").Enabled = true;
                    oForm.Items.Item("Mat1").Enabled = true;
                    oForm.Items.Item("Btn_Apply").Enabled = true;
                    oForm.Items.Item("Btn_Cancel").Enabled = false;
                    
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //?ëÏÜç?êÏóê ?∞Î•∏ Í∂åÌïúÎ≥? ?¨ÏóÖ?? ÏΩ§Î≥¥Î∞ïÏä§?∏ÌåÖ

                    oForm.EnableMenu("1281", true); //Î¨∏ÏÑúÏ∞æÍ∏∞
                    oForm.EnableMenu("1282", false); //Î¨∏ÏÑúÏ∂îÍ?
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = false;

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //?ëÏÜç?êÏóê ?∞Î•∏ Í∂åÌïúÎ≥? ?¨ÏóÖ?? ÏΩ§Î≥¥Î∞ïÏä§?∏ÌåÖ

                    oForm.EnableMenu("1281", false); //Î¨∏ÏÑúÏ∞æÍ∏∞
                    oForm.EnableMenu("1282", true); //Î¨∏ÏÑúÏ∂îÍ?
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("YM").Enabled = false;
                    oForm.Items.Item("Comments").Enabled = false;

                    if (oForm.Items.Item("StatYN").Specific.Value == "Y")
                    {
                        oForm.Items.Item("FieldCo").Enabled = false;
                        oForm.Items.Item("Mat1").Enabled = false;
                        oForm.Items.Item("Btn_Apply").Enabled = false;
                        oForm.Items.Item("Btn_Cancel").Enabled = true;
                    }
                    else
                    {
                        oForm.Items.Item("FieldCo").Enabled = true;
                        oForm.Items.Item("Mat1").Enabled = true;
                        oForm.Items.Item("Btn_Apply").Enabled = true;
                        oForm.Items.Item("Btn_Cancel").Enabled = false;
                    }

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false); //?ëÏÜç?êÏóê ?∞Î•∏ Í∂åÌïúÎ≥? ?¨ÏóÖ?? ÏΩ§Î≥¥Î∞ïÏä§?∏ÌåÖ

                    oForm.EnableMenu("1281", true); //Î¨∏ÏÑúÏ∞æÍ∏∞
                    oForm.EnableMenu("1282", true); //Î¨∏ÏÑúÏ∂îÍ?
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY124_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Îß§Ìä∏Î¶?ä§ ?? Ï∂îÍ?
        /// </summary>
        private void PH_PY124_AddMatrixRow()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY124B.GetValue("U_Seq", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY124B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY124B.InsertRecord(oRow);
                        }
                        oDS_PH_PY124B.Offset = oRow;
                        oDS_PH_PY124B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY124B.SetValue("U_Seq", oRow, "");
                        oDS_PH_PY124B.SetValue("U_MSTCOD", oRow, "");
                        oDS_PH_PY124B.SetValue("U_MSTNAM", oRow, "");
                        oDS_PH_PY124B.SetValue("U_BeneAmt", oRow, "0");
                        oDS_PH_PY124B.SetValue("U_BillAmt", oRow, "0");
                        oDS_PH_PY124B.SetValue("U_CardAmt", oRow, "0");
                        oDS_PH_PY124B.SetValue("U_TotAmt", oRow, "0");
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY124B.Offset = oRow - 1;
                        oDS_PH_PY124B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY124B.SetValue("U_Seq", oRow, "");
                        oDS_PH_PY124B.SetValue("U_MSTCOD", oRow, "");
                        oDS_PH_PY124B.SetValue("U_MSTNAM", oRow, "");
                        oDS_PH_PY124B.SetValue("U_BeneAmt", oRow, "0");
                        oDS_PH_PY124B.SetValue("U_BillAmt", oRow, "0");
                        oDS_PH_PY124B.SetValue("U_CardAmt", oRow, "0");
                        oDS_PH_PY124B.SetValue("U_TotAmt", oRow, "0");
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY124B.Offset = oRow;
                    oDS_PH_PY124B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY124B.SetValue("U_Seq", oRow, "");
                    oDS_PH_PY124B.SetValue("U_MSTCOD", oRow, "");
                    oDS_PH_PY124B.SetValue("U_MSTNAM", oRow, "");
                    oDS_PH_PY124B.SetValue("U_BeneAmt", oRow, "0");
                    oDS_PH_PY124B.SetValue("U_BillAmt", oRow, "0");
                    oDS_PH_PY124B.SetValue("U_CardAmt", oRow, "0");
                    oDS_PH_PY124B.SetValue("U_TotAmt", oRow, "0");
                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY124_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        private bool PH_PY124_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oDS_PH_PY124A.GetValue("U_CLTCOD", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("?¨ÏóÖ?•Ï? ?ÑÏàò?ÖÎãà??.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                //?ÅÏö©?úÏûë??
                if (string.IsNullOrEmpty(oDS_PH_PY124A.GetValue("U_YM", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("?ÅÏö©?úÏûë?îÏ? ?ÑÏàò?ÖÎãà??.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                //Code & Name ?ùÏÑ±
                oDS_PH_PY124A.SetValue("Code", 0, oDS_PH_PY124A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY124A.GetValue("U_YM", 0).ToString().Trim());
                oDS_PH_PY124A.SetValue("NAME", 0, oDS_PH_PY124A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY124A.GetValue("U_YM", 0).ToString().Trim());

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (!string.IsNullOrEmpty(dataHelpClass.Get_ReData("Code", "Code", "[@PH_PY124A]", "'" + oDS_PH_PY124A.GetValue("Code", 0).ToString().Trim() + "'", "")))
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("?¥Î? Ï°¥Ïû¨?òÎäî ÏΩîÎìú?ÖÎãà??.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        return functionReturnValue;
                    }
                }

                //?ºÏù∏
                if (oMat1.VisualRowCount >= 1)
                {
                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                    {

                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("?ºÏù∏ ?∞Ïù¥?∞Í? ?ÜÏäµ?àÎã§.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return functionReturnValue;
                }

                oMat1.FlushToDataSource();

                //Matrix ÎßàÏ?Îß? ?? ??†ú(DB ?Ä?•Ïãú)
                if (oDS_PH_PY124B.Size > 1)
                {
                    oDS_PH_PY124B.RemoveRecord(oDS_PH_PY124B.Size - 1);
                }

                oMat1.LoadFromDataSource();
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY124_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY124_MTX01
        /// </summary>
        private void PH_PY124_MTX01()
        {
            int i;
            string sQry;
            int ErrNum = 0;

            string Param01;
            string Param02;
            string Param03;
            string Param04;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("Param01").Specific.Value;
                Param02 = oForm.Items.Item("Param01").Specific.Value;
                Param03 = oForm.Items.Item("Param01").Specific.Value;
                Param04 = oForm.Items.Item("Param01").Specific.Value;

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                sQry = "SELECT 10";
                oRecordSet.DoQuery(sQry);

                oMat1.Clear();
                oMat1.FlushToDataSource();
                oMat1.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PH_PY124B.InsertRecord((i));
                    }
                    oDS_PH_PY124B.Offset = i;
                    oDS_PH_PY124B.SetValue("U_COL01", i, oRecordSet.Fields.Item(0).Value);
                    oDS_PH_PY124B.SetValue("U_COL02", i, oRecordSet.Fields.Item(1).Value);
                    oRecordSet.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "Í±? Ï°∞ÌöåÏ§?...!";
                }
                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();
                oForm.Update();

                ProgressBar01.Stop();
            }
            catch (Exception ex)
            {
                ProgressBar01.Stop();
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("Í≤∞Í≥ºÍ∞Ä Ï°¥Ïû¨?òÏ? ?äÏäµ?àÎã§.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY001_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY124_FormClear
        /// </summary>
        private void PH_PY124_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = DataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY124'", "");
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY124_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY124_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            short ErrNumm = 0;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY124A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    ErrNumm = 1;
                    throw new Exception();
                }
                if (ValidateType == "?òÏ†ï")
                {

                }
                else if (ValidateType == "?âÏÇ≠??")
                {

                }
                else if (ValidateType == "Ï∑®ÏÜå")
                {

                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (ErrNumm == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("?¥ÎãπÎ¨∏ÏÑú?? ?§Î•∏?¨Ïö©?êÏóê ?òÌï¥ Ï∑®ÏÜå?òÏóà?µÎãà??. ?ëÏóÖ?? ÏßÑÌñâ?†Ïàò ?ÜÏäµ?àÎã§.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// ?ëÏ? ?åÏùº ?ÖÎ°ú??
        /// </summary>
        [STAThread]
        private void PH_PY124_Excel_Upload()
        {
            int loopCount;
            int j;
            int CheckLine;
            int i;
            bool sucessFlag = false;
            short columnCount = 7; //?ëÏ? Ïª¨Îüº??
            short columnCount2 = 7; //?ëÏ? Ïª¨Îüº??
            string sFile;
            double TOTCNT;
            int V_StatusCnt;
            int oProValue;
            int tRow;
            //bool CheckYN = true;

            CommonOpenFileDialog commonOpenFileDialog = new CommonOpenFileDialog();

            commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("Excel Files", "*.xls;*.xlsx"));
            commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("Î™®Îì† ?åÏùº", "*.*"));
            commonOpenFileDialog.IsFolderPicker = false;

            if (commonOpenFileDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                sFile = commonOpenFileDialog.FileName;
            }
            else //Cancel Î≤ÑÌäº ?¥Î¶≠
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

            //?ëÏ? Object ?∞Í≤∞
            //?îÏãú?? Í∞ùÏ≤¥Ï∞∏Ï°∞ ?? Excel.exe Î©îÎ™®Î¶? Î∞òÌôò?? ?àÎê®, ?ÑÎûò?Ä Í∞ôÏù¥ Î™ÖÏãú?? Ï∞∏Ï°∞Î°? ?†Ïñ∏
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
                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("?úÏûë!", xlRow.Count, false);
                Microsoft.Office.Interop.Excel.Range[] t = new Microsoft.Office.Interop.Excel.Range[columnCount2 + 1];
                for (loopCount = 1; loopCount <= columnCount2; loopCount++)
                {
                    t[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[1, loopCount];
                }

                // Ï≤? ?Ä?¥Ì? ÎπÑÍµê
                if (Convert.ToString(t[1].Value) != "?ºÎ†®Î≤àÌò∏")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("A?? Ï≤´Î≤àÏß? ?? ?Ä?¥Ì??Ä ?ºÎ†®Î≤àÌò∏", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[2].Value) != "?¨Î≤à")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("B?? ?êÎ≤àÏß? ?? ?Ä?¥Ì??Ä ?¨Î≤à", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[3].Value) != "?¥Î¶Ñ")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("C?? ?∏Î≤àÏß? ?? ?Ä?¥Ì??Ä ?¥Î¶Ñ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[4].Value) != "Î≤†ÎÑ§?ºÏïÑ")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("D?? ?∏Î≤àÏß? ?? ?Ä?¥Ì??Ä Î≤†ÎÑ§?ºÏïÑ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[5].Value) != "?ÅÏàòÏ¶?")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("E?? ?∏Î≤àÏß? ?? ?Ä?¥Ì??Ä ?ÅÏàòÏ¶?", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[6].Value) != "Î≥µÏ?Ïπ¥Îìú(Íµ?Ç¥)")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("F?? ?∏Î≤àÏß? ?? ?Ä?¥Ì??Ä Î≥µÏ?Ïπ¥Îìú(Íµ?Ç¥)", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[7].Value) != "Ï¥ùÌï©Í≥?(??)")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("G?? ?∏Î≤àÏß? ?? ?Ä?¥Ì??Ä Ï¥ùÌï©Í≥?(??)", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }

                //?ÑÎ°úÍ∑∏Î†à?? Î∞?    ///////////////////////////////////////
                ProgressBar01.Text = "?∞Ïù¥?? ?ΩÎäîÏ§?...!";

                //ÏµúÎ?Í∞? Íµ¨ÌïòÍ∏? ///////////////////////////////////////
                TOTCNT = xlsh.UsedRange.Rows.Count;

                V_StatusCnt = Convert.ToInt32(Math.Round(TOTCNT / 50, 0));
                oProValue = 1;
                tRow = 1;
                /////////////////////////////////////////////////////

                for (i = 2; i <= xlsh.UsedRange.Rows.Count; i++)
                {
                    Microsoft.Office.Interop.Excel.Range[] r = new Microsoft.Office.Interop.Excel.Range[columnCount + 1];

                    for (loopCount = 1; loopCount <= columnCount; loopCount++)
                    {
                        r[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[i, loopCount];
                    }
                    //CheckYN = false;
                    for (j = 0; j <= oDS_PH_PY124B.Size - 1; j++)
                    {

                        if (Convert.ToString(r[1].Value) == oDS_PH_PY124B.GetValue("U_MSTCOD", j).ToString().Trim())
                        {
                            //CheckYN = true;
                            CheckLine = j;
                        }
                    }

                    //ÎßàÏ?ÎßâÌñâ ?úÍ±∞
                    if (string.IsNullOrEmpty(oDS_PH_PY124B.GetValue("U_MSTCOD", oDS_PH_PY124B.Size - 1).ToString().Trim()))
                    {
                        oDS_PH_PY124B.RemoveRecord(oDS_PH_PY124B.Size - 1);
                    }

                    oDS_PH_PY124B.InsertRecord(oDS_PH_PY124B.Size);
                    oDS_PH_PY124B.Offset = oDS_PH_PY124B.Size - 1;
                    oDS_PH_PY124B.SetValue("U_LineNum", oDS_PH_PY124B.Size - 1, Convert.ToString(r[1].Value));
                    oDS_PH_PY124B.SetValue("U_Seq", oDS_PH_PY124B.Size - 1, Convert.ToString(r[1].Value));
                    oDS_PH_PY124B.SetValue("U_MSTCOD", oDS_PH_PY124B.Size - 1, Convert.ToString(r[2].Value));
                    oDS_PH_PY124B.SetValue("U_MSTNAM", oDS_PH_PY124B.Size - 1, Convert.ToString(r[3].Value));
                    oDS_PH_PY124B.SetValue("U_BeneAmt", oDS_PH_PY124B.Size - 1, Convert.ToString(r[4].Value));
                    oDS_PH_PY124B.SetValue("U_BillAmt", oDS_PH_PY124B.Size - 1, Convert.ToString(r[5].Value));
                    oDS_PH_PY124B.SetValue("U_CardAmt", oDS_PH_PY124B.Size - 1, Convert.ToString(r[6].Value));
                    oDS_PH_PY124B.SetValue("U_TotAmt", oDS_PH_PY124B.Size - 1, Convert.ToString(r[7].Value));

                    if ((TOTCNT > 50 && tRow == oProValue * V_StatusCnt) || TOTCNT <= 50)
                    {
                        ProgressBar01.Text = tRow + "/ " + TOTCNT + " Í±? Ï≤òÎ¶¨Ï§?...!";
                        ProgressBar01.Value += 1;
                    }
                    tRow += 1;
                }
                ProgressBar01.Value += 1;
                ProgressBar01.Text = ProgressBar01.Value + "/" + (xlRow.Count - 1) + "Í±? Loding...!";

                //?ºÏù∏Î≤àÌò∏ ?¨Ï†ï??
                for (i = 0; i <= oDS_PH_PY124B.Size - 1; i++)
                {
                    oDS_PH_PY124B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                }
                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY124_Excel_Upload:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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

                ProgressBar01.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);

                if (sucessFlag == true)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("?ëÏ? Loding ?ÑÎ£å", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY124_DataApply
        /// </summary>
        /// <param name="CLTCOD"></param>
        /// <param name="YM"></param>
        /// <returns></returns>
        private bool PH_PY124_DataApply(string CLTCOD, string YM)
        {
            bool functionReturnValue = false;
            string sQry;
            string AMTLen;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                oMat1.FlushToDataSource();

                if (oForm.Items.Item("FieldCo").Specific.Value.ToString().Trim().Length == 1)
                {
                    AMTLen = Convert.ToString(Convert.ToDouble("0") + oForm.Items.Item("FieldCo").Specific.Value).ToString().Trim();
                }
                else
                {
                    AMTLen = oForm.Items.Item("FieldCo").Specific.Value.ToString().Trim();
                }

                sQry = "";
                sQry += " update [@PH_PY109B]";
                sQry += " set U_AMT" + AMTLen + "=isnull(U_AMT" + AMTLen + ",0)  + isnull(b.U_TotAmt,0)";
                sQry += " from [@PH_PY109B] a left join [@PH_PY124B] b on left(a.code,1) = left(b.code,1) and SUBSTRING(a.code,2,4) = right(b.code,4) and a.U_MSTCOD  = b.U_MSTCOD";
                sQry += " where a.code ='" + CLTCOD + codeHelpClass.Right(YM, 4) + "111'";

                oRecordSet.DoQuery(sQry);

                sQry = "";
                sQry += " update [@PH_PY124A] set U_statYN = 'Y' where U_NaviDoc ='" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + oForm.Items.Item("YM").Specific.Value.ToString().Trim() + "'";

                oRecordSet.DoQuery(sQry);

                oForm.Items.Item("StatYN").Specific.Value = "Y";
                oForm.Items.Item("Test").Click((SAPbouiCOM.BoCellClickType.ct_Regular));

                oForm.Items.Item("FieldCo").Enabled = false;
                oForm.Items.Item("Mat1").Enabled = false;
                oForm.Items.Item("Btn_Apply").Enabled = false;
                oForm.Items.Item("Btn_Cancel").Enabled = true;

                PSH_Globals.SBO_Application.StatusBar.SetText("Í∏âÏÉÅ?¨Î??? ?êÎ£å?? Í∏àÏï°?? ?ÅÏö© ?òÏóà?µÎãà??.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
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
        /// PH_PY124_DataCancel
        /// </summary>
        /// <param name="CLTCOD"></param>
        /// <param name="YM"></param>
        /// <returns></returns>
        private bool PH_PY124_DataCancel(string CLTCOD, string YM)
        {
            bool functionReturnValue = false;
            string sQry;
            string AMTLen;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                oMat1.FlushToDataSource();

                if (oForm.Items.Item("FieldCo").Specific.Value.ToString().Trim().Length == 1)
                {
                    AMTLen = Convert.ToString(Convert.ToDouble("0") + oForm.Items.Item("FieldCo").Specific.Value).ToString().Trim();
                }
                else
                {
                    AMTLen = oForm.Items.Item("FieldCo").Specific.Value.ToString().Trim();
                }

                sQry = "";
                sQry += " update [@PH_PY109B]";
                sQry += " set U_AMT" + AMTLen + "=isnull(U_AMT" + AMTLen + ",0)  - isnull(b.U_TotAmt,0)";
                sQry += " from [@PH_PY109B] a left join [@PH_PY124B] b on left(a.code,1) = left(b.code,1) and SUBSTRING(a.code,2,4) = right(b.code,4) and a.U_MSTCOD  = b.U_MSTCOD";
                sQry += " where a.code ='" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + codeHelpClass.Right(oForm.Items.Item("YM").Specific.Value.ToString().Trim(), 4) + "111'";

                oRecordSet.DoQuery(sQry);

                sQry = "";
                sQry += " update [@PH_PY124A] set U_statYN = 'N' where U_NaviDoc ='" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + oForm.Items.Item("YM").Specific.Value.ToString().Trim() + "'";

                oRecordSet.DoQuery(sQry);

                oForm.Items.Item("StatYN").Specific.Value = "N";
                oForm.Items.Item("FieldCo").Enabled = true;
                oForm.Items.Item("Mat1").Enabled = true;
                oForm.Items.Item("Btn_Apply").Enabled = true;
                oForm.Items.Item("Btn_Cancel").Enabled = false;

                PSH_Globals.SBO_Application.StatusBar.SetText("Í∏âÏÉÅ?¨Î??? ?êÎ£å?? Í∏àÏï°?? ?ÅÏö© ?òÏóà?µÎãà??.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

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
        /// ITEM_PRESSED ?¥Î≤§??
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent Í∞ùÏ≤¥</param>
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
                            if (PH_PY124_DataValidCheck() == false)
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
                                PH_PY124_DataApply(CLTCOD, YM);
                                CheckDataApply = false;
                            }
                            PH_PY124_FormItemEnabled();
                            PH_PY124_AddMatrixRow();
                        }
                    }
                    if (pVal.ItemUID == "Btn_UPLOAD")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY124_Excel_Upload);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start(); 
                        PH_PY124_AddMatrixRow();
                    }
                    if (pVal.ItemUID == "Btn_Cancel")
                    {
                        PH_PY124_DataCancel(CLTCOD, YM);
                    }
                    if (pVal.ItemUID == "Btn_Apply")
                    {
                        CLTCOD = oDS_PH_PY124A.GetValue("U_CLTCOD", 0).ToString().Trim();
                        YM = oDS_PH_PY124A.GetValue("U_YM", 0).ToString().Trim();
                        if (oMat1.RowCount > 1)
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                CheckDataApply = true;
                                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                PH_PY124_DataApply(CLTCOD, YM);
                            }
                        }
                        else
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("Î≤†ÎÑ§?ºÏïÑ ?êÎ£åÍ∞Ä ?ÜÏäµ?àÎã§.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        /// Raise_EVENT_GOT_FOCUS ?¥Î≤§??
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent Í∞ùÏ≤¥</param>
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
        /// Raise_EVENT_CLICK ?¥Î≤§??
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent Í∞ùÏ≤¥</param>
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
            string sQry;
            string FieldCo;
            int ErrCode = 0;
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

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ItemUID == "YM")
                            {

                                CLTCOD = oDS_PH_PY124A.GetValue("U_CLTCOD", 0).ToString().Trim();
                                YM = codeHelpClass.Right(oDS_PH_PY124A.GetValue("U_YM", 0).ToString().Trim(), 4);

                                if (!string.IsNullOrEmpty(oDS_PH_PY124A.GetValue("U_FieldCo", 0).ToString().Trim()))
                                {
                                    FieldCo = " = '" + oDS_PH_PY124A.GetValue("U_FieldCo", 0).ToString().Trim();
                                }
                                else
                                {
                                    FieldCo = " like '%";
                                }
                                sQry = "select U_Sequence from [@PH_PY109Z] where code ='" + CLTCOD + YM + "111'";
                                oRecordSet.DoQuery(sQry);
                                if (oRecordSet.RecordCount == 0)
                                {
                                    ErrCode = 1;
                                    throw new Exception();
                                }
                                else
                                {
                                    sQry = "select distinct U_Sequence,U_PDName from [@PH_PY109Z] where code ='" + CLTCOD + YM + "111' and u_sequence" + FieldCo + "' order by 1";
                                    dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("FieldCo").Specific, "");
                                    oForm.Items.Item("FieldCo").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                    oForm.Items.Item("FieldCo").DisplayDesc = true;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if(ErrCode == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("Í∏âÏÉÅ?¨Î??ôÏûêÎ£? ?ÖÎ†•?Ä ?ÑÏàò?ÖÎãà??.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //Î©îÎ™®Î¶? ?¥Ï†ú
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
                    PH_PY124_FormItemEnabled();
                    PH_PY124_AddMatrixRow();
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
        /// FORM_UNLOAD ?¥Î≤§??
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent Í∞ùÏ≤¥</param>
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY124A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY124B);
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
        /// ROW_DELETE(Raise_FormMenuEvent?êÏÑú ?∏Ï∂ú), ?¥Îãπ ?¥Îûò?§Ïóê?úÎäî ?¨Ïö©?òÏ? ?äÏùå
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
                            if (PSH_Globals.SBO_Application.MessageBox("?ÑÏû¨ ?îÎ©¥?¥Ïö©?ÑÏ≤¥Î•? ?úÍ±∞ ?òÏãúÍ≤†Ïäµ?àÍπå? Î≥µÍµ¨?? ?? ?ÜÏäµ?àÎã§.", 2, "Yes", "No") == 2)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1284":
                            CLTCOD = oDS_PH_PY124A.GetValue("U_CLTCOD", 0).ToString().Trim();
                            YM = oDS_PH_PY124A.GetValue("U_YM", 0).ToString().Trim();
                            PH_PY124_DataCancel(CLTCOD, YM);
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
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY124A", "Code"); //?ëÏÜç?? Í∂åÌïú?? ?∞Î•∏ ?¨ÏóÖ?? Î≥¥Í∏∞
                            PH_PY124_FormItemEnabled();
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY124_FormItemEnabled();
                            PH_PY124_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //Î¨∏ÏÑúÏ∞æÍ∏∞
                            PH_PY124_FormItemEnabled();
                            PH_PY124_AddMatrixRow();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //Î¨∏ÏÑúÏ∂îÍ?
                            PH_PY124_FormItemEnabled();
                            PH_PY124_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY124_FormItemEnabled();
                            CLTCOD = oDS_PH_PY124A.GetValue("U_CLTCOD", 0).ToString().Trim();
                            YM = codeHelpClass.Right(oDS_PH_PY124A.GetValue("U_YM", 0).ToString().Trim(), 4);
                            break;
                        case "1293": //?âÏÇ≠??
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat1, oDS_PH_PY124B, "U_JIGCOD");
                            PH_PY124_AddMatrixRow();
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

//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE Í∞úÏ≤¥?? Í∏∞Î≥∏ ?çÏÑ±?? ?ïÏù∏?? ?? ?ÜÏäµ?àÎã§. ?êÏÑ∏?? ?¥Ïö©?Ä ?§Ïùå?? Ï∞∏Ï°∞?òÏã≠?úÏò§. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("StatYN").Specific.VALUE = "Y";

//            oForm.Items.Item("Test").Click((SAPbouiCOM.BoCellClickType.ct_Regular));

//            oForm.Items.Item("FieldCo").Enabled = false;
//            oForm.Items.Item("Mat1").Enabled = false;
//            oForm.Items.Item("Btn_Apply").Enabled = false;
//            oForm.Items.Item("Btn_Cancel").Enabled = true;

//            MDC_Globals.Sbo_Application.StatusBar.SetText("Í∏âÏÉÅ?¨Î??? ?êÎ£å?? Í∏àÏï°?? ?ÅÏö© ?òÏóà?µÎãà??.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//            //UPGRADE_NOTE: oRecordSet Í∞úÏ≤¥?? Í∞ÄÎπÑÏ?Í∞Ä ?òÏßë?òÏñ¥?? ?åÎ©∏?©Îãà??. ?êÏÑ∏?? ?¥Ïö©?Ä ?§Ïùå?? Ï∞∏Ï°∞?òÏã≠?úÏò§. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            return functionReturnValue;
//        PH_PY124_DataApply_Error:

//            //UPGRADE_NOTE: oRecordSet Í∞úÏ≤¥?? Í∞ÄÎπÑÏ?Í∞Ä ?òÏßë?òÏñ¥?? ?åÎ©∏?©Îãà??. ?êÏÑ∏?? ?¥Ïö©?Ä ?§Ïùå?? Ï∞∏Ï°∞?òÏã≠?úÏò§. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY124_DataApply_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }

//        private bool PH_PY124_DataCancel(ref string CLTCOD, ref string YM)
//        {
//            bool functionReturnValue = false;
//            string sQry = null;
//            SAPbobsCOM.Recordset oRecordSet = null;
//            string Tablename = null;
//            string sTablename = null;
//            string AMTLen = null;


//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            // ERROR: Not supported in C#: OnErrorStatement


//            functionReturnValue = false;

//            oMat1.FlushToDataSource();

//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE Í∞úÏ≤¥?? Í∏∞Î≥∏ ?çÏÑ±?? ?ïÏù∏?? ?? ?ÜÏäµ?àÎã§. ?êÏÑ∏?? ?¥Ïö©?Ä ?§Ïùå?? Ï∞∏Ï°∞?òÏã≠?úÏò§. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            if (Strings.Len(Strings.Trim(oForm.Items.Item("FieldCo").Specific.VALUE)) == 1)
//            {
//                //UPGRADE_WARNING: oForm.Items().Specific.VALUE Í∞úÏ≤¥?? Í∏∞Î≥∏ ?çÏÑ±?? ?ïÏù∏?? ?? ?ÜÏäµ?àÎã§. ?êÏÑ∏?? ?¥Ïö©?Ä ?§Ïùå?? Ï∞∏Ï°∞?òÏã≠?úÏò§. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                AMTLen = Strings.Trim(Convert.ToString(Convert.ToDouble("0") + oForm.Items.Item("FieldCo").Specific.VALUE));
//            }
//            else
//            {
//                //UPGRADE_WARNING: oForm.Items().Specific.VALUE Í∞úÏ≤¥?? Í∏∞Î≥∏ ?çÏÑ±?? ?ïÏù∏?? ?? ?ÜÏäµ?àÎã§. ?êÏÑ∏?? ?¥Ïö©?Ä ?§Ïùå?? Ï∞∏Ï°∞?òÏã≠?úÏò§. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                AMTLen = Strings.Trim(oForm.Items.Item("FieldCo").Specific.VALUE);
//            }


//            sQry = "";
//            sQry = sQry + " update [@PH_PY109B]";
//            sQry = sQry + " set U_AMT" + AMTLen + "=isnull(U_AMT" + AMTLen + ",0)  - isnull(b.U_TotAmt,0)";
//            sQry = sQry + " from [@PH_PY109B] a left join [@PH_PY124B] b on left(a.code,1) = left(b.code,1) and SUBSTRING(a.code,2,4) = right(b.code,4) and a.U_MSTCOD  = b.U_MSTCOD";
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE Í∞úÏ≤¥?? Í∏∞Î≥∏ ?çÏÑ±?? ?ïÏù∏?? ?? ?ÜÏäµ?àÎã§. ?êÏÑ∏?? ?¥Ïö©?Ä ?§Ïùå?? Ï∞∏Ï°∞?òÏã≠?úÏò§. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            sQry = sQry + " where a.code ='" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + Strings.Right(Strings.Trim(oForm.Items.Item("YM").Specific.VALUE), 4) + "111'";

//            oRecordSet.DoQuery(sQry);

//            sQry = "";
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE Í∞úÏ≤¥?? Í∏∞Î≥∏ ?çÏÑ±?? ?ïÏù∏?? ?? ?ÜÏäµ?àÎã§. ?êÏÑ∏?? ?¥Ïö©?Ä ?§Ïùå?? Ï∞∏Ï°∞?òÏã≠?úÏò§. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            sQry = sQry + " update [@PH_PY124A] set U_statYN = 'N' where U_NaviDoc ='" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + Strings.Trim(oForm.Items.Item("YM").Specific.VALUE) + "'";

//            oRecordSet.DoQuery(sQry);

//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE Í∞úÏ≤¥?? Í∏∞Î≥∏ ?çÏÑ±?? ?ïÏù∏?? ?? ?ÜÏäµ?àÎã§. ?êÏÑ∏?? ?¥Ïö©?Ä ?§Ïùå?? Ï∞∏Ï°∞?òÏã≠?úÏò§. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("StatYN").Specific.VALUE = "N";
//            oForm.Items.Item("FieldCo").Enabled = true;
//            oForm.Items.Item("Mat1").Enabled = true;
//            oForm.Items.Item("Btn_Apply").Enabled = true;
//            oForm.Items.Item("Btn_Cancel").Enabled = false;

//            MDC_Globals.Sbo_Application.StatusBar.SetText("Í∏âÏÉÅ?¨Î??? ?êÎ£å?? Í∏àÏï°?? ?ÅÏö© ?òÏóà?µÎãà??.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//            //UPGRADE_NOTE: oRecordSet Í∞úÏ≤¥?? Í∞ÄÎπÑÏ?Í∞Ä ?òÏßë?òÏñ¥?? ?åÎ©∏?©Îãà??. ?êÏÑ∏?? ?¥Ïö©?Ä ?§Ïùå?? Ï∞∏Ï°∞?òÏã≠?úÏò§. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            return functionReturnValue;
//        PH_PY124_DataCancel_Error:

//            //UPGRADE_NOTE: oRecordSet Í∞úÏ≤¥?? Í∞ÄÎπÑÏ?Í∞Ä ?òÏßë?òÏñ¥?? ?åÎ©∏?©Îãà??. ?êÏÑ∏?? ?¥Ïö©?Ä ?§Ïùå?? Ï∞∏Ï°∞?òÏã≠?úÏò§. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY124_DataCancel_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }
//    }
//}
