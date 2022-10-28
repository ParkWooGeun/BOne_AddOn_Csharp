using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// ?úÌíà ?êÏû¨Î£? Î≥Ä??
    /// </summary>
    internal class PS_CO800 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_CO800H; //?±Î°ù?§Îçî
        private SAPbouiCOM.DBDataSource oDS_PS_CO800L; //?±Î°ù?ºÏù∏
        private string oLastItemUID01; //?¥Îûò?§Ïóê?? ?†ÌÉù?? ÎßàÏ?Îß? ?ÑÏù¥?? UidÍ∞?
        private string oLastColUID01; //ÎßàÏ?ÎßâÏïÑ?¥ÌÖú?? Î©îÌä∏Î¶?ä§?ºÍ≤Ω?∞Ïóê ÎßàÏ?Îß? ?†ÌÉù?? Col?? UidÍ∞?
        private int oLastColRow01; //ÎßàÏ?ÎßâÏïÑ?¥ÌÖú?? Î©îÌä∏Î¶?ä§?ºÍ≤Ω?∞Ïóê ÎßàÏ?Îß? ?†ÌÉù?? RowÍ∞?
        private int oSeq;

        /// <summary>
        /// Form ?∏Ï∂ú
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO800.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_CO800_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_CO800");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocNum";

                oForm.Freeze(true);
                oForm.EnableMenu("1293", true);

                PS_CO800_CreateItems();
                PS_CO800_Initial_Setting();
                PS_CO800_FormItemEnabled();
                PS_CO800_FormClear();
                PS_CO800_AddMatrixRow(0, oMat01.RowCount);

                oForm.EnableMenu("1283", false); // ??†ú
                oForm.EnableMenu("1286", false); // ?´Í∏∞
                oForm.EnableMenu("1287", false); // Î≥µÏ†ú
                oForm.EnableMenu("1284", true); // Ï∑®ÏÜå
                oForm.EnableMenu("1293", true); // ?âÏÇ≠??
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //Î©îÎ™®Î¶? ?¥Ï†ú
            }
        }

        /// <summary>
        /// ?îÎ©¥ Item ?ùÏÑ±
        /// </summary>
        private void PS_CO800_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oDS_PS_CO800H = oForm.DataSources.DBDataSources.Item("@PS_CO800H");
                oDS_PS_CO800L = oForm.DataSources.DBDataSources.Item("@PS_CO800L");

                oMat01 = oForm.Items.Item("Mat01").Specific; //Îß§Ìä∏Î¶?ä§ ?∞Ïù¥?? ??
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;

                //?¨ÏóÖ?? Î¶¨Ïä§??
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";  
                oRecordSet01.DoQuery(sQry); 

                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// Initial Setting
        /// </summary>
        private void PS_CO800_Initial_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); //?¨ÏóÖ??
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Î™®Îìú?? ?∞Î•∏ ?ÑÏù¥?? ?§Ï†ï
        /// </summary>
        private void PS_CO800_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("MstCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oMat01.Columns.Item("ItemCode").Editable = true;
                    oMat01.Columns.Item("BatchNum").Editable = true;
                    oMat01.Columns.Item("WhsCode").Editable = true;
                    oMat01.Columns.Item("RitmCode").Editable = true;
                    oMat01.Columns.Item("MoveQty").Editable = true;
                    oForm.Items.Item("DoNumber").Enabled = true;

                    if (oForm.Items.Item("OKYN").Specific.Value == "N")
                    {
                        if (oForm.Items.Item("Status").Specific.Value == "C")
                        {
                            oForm.Items.Item("Btn02").Enabled = false;
                            oForm.Items.Item("Btn03").Enabled = false;
                        }
                        else
                        {
                            oForm.Items.Item("Btn02").Enabled = true;
                            oForm.Items.Item("Btn03").Enabled = true;
                        }
                    }
                    else
                    {
                        oForm.Items.Item("Btn02").Enabled = false;
                        oForm.Items.Item("Btn03").Enabled = false;
                    }
                    if (oForm.Items.Item("ChFwYN").Specific.Value == "Y")
                    {
                        oForm.Items.Item("DoNumber").Enabled = false;
                        oForm.Items.Item("Btn02").Enabled = false;
                    }
                    if (oForm.Items.Item("ChRvYN").Specific.Value == "Y" || oForm.Items.Item("ChFwYN").Specific.Value == "")
                    {
                        oForm.Items.Item("Btn03").Enabled = false;
                    }
                    dataHelpClass.CLTCOD_Select(oForm, "BPLId", true); 
                    oDS_PS_CO800H.SetValue("U_Docdate", 0, DateTime.Now.ToString("yyyyMMdd"));
                    oDS_PS_CO800H.SetValue("U_MstCode", 0, dataHelpClass.User_MSTCOD());
                    oDS_PS_CO800H.SetValue("U_MstName", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + dataHelpClass.User_MSTCOD() + "'", ""));
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("DoNumber").Enabled = true;
                    

                    if (oForm.Items.Item("OKYN").Specific.Value == "N")
                    {
                        if (oForm.Items.Item("Status").Specific.Value == "C")
                        {
                            oForm.Items.Item("Btn02").Enabled = false;
                            oForm.Items.Item("Btn03").Enabled = false;
                        }
                        else
                        {
                            oForm.Items.Item("Btn02").Enabled = true;
                            oForm.Items.Item("Btn03").Enabled = true;
                        }
                    }
                    else
                    {
                        oForm.Items.Item("Btn02").Enabled = false;
                        oForm.Items.Item("Btn03").Enabled = false;
                    }
                    if (oForm.Items.Item("ChFwYN").Specific.Value == "Y")
                    {
                        oForm.Items.Item("Btn02").Enabled = false;
                    }
                    if (oForm.Items.Item("ChRvYN").Specific.Value == "Y" || oForm.Items.Item("ChFwYN").Specific.Value == "")
                    {
                        oForm.Items.Item("Btn03").Enabled = false;
                    }
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    if (oForm.Items.Item("ChFwYN").Specific.Value == "Y")
                    {
                        oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("MstCode").Enabled = false;
                        oForm.Items.Item("DocDate").Enabled = false;
                        oForm.Items.Item("BPLId").Enabled = false;
                        oMat01.Columns.Item("ItemCode").Editable = false;
                        oMat01.Columns.Item("BatchNum").Editable = false;
                        oMat01.Columns.Item("WhsCode").Editable = false;
                        oMat01.Columns.Item("RitmCode").Editable = false;
                        oMat01.Columns.Item("MoveQty").Editable = false;
                        oForm.Items.Item("DoNumber").Enabled = false;
                    }
                    if (oForm.Items.Item("OKYN").Specific.Value == "N")
                    {
                        oForm.Items.Item("MstCode").Enabled = false;
                        oForm.Items.Item("DocDate").Enabled = false;
                        oForm.Items.Item("BPLId").Enabled = false;
                        oMat01.Columns.Item("ItemCode").Editable = false;
                        oMat01.Columns.Item("BatchNum").Editable = false;
                        oMat01.Columns.Item("WhsCode").Editable = false;
                        oMat01.Columns.Item("RitmCode").Editable = false;
                        oMat01.Columns.Item("MoveQty").Editable = false;
                        if (oForm.Items.Item("Status").Specific.Value == "C")
                        {
                            oForm.Items.Item("Btn02").Enabled = false;
                            oForm.Items.Item("Btn03").Enabled = false;
                        }
                        else
                        {
                            oForm.Items.Item("Btn02").Enabled = true;
                            oForm.Items.Item("Btn03").Enabled = true;
                        }
                    }
                    else
                    {
                        oForm.Items.Item("Btn02").Enabled = false;
                        oForm.Items.Item("Btn03").Enabled = false;
                    }
                    if (oForm.Items.Item("ChFwYN").Specific.Value == "Y")
                    {
                        oForm.Items.Item("Btn02").Enabled = false;
                    }
                    if (oForm.Items.Item("ChRvYN").Specific.Value == "Y" || oForm.Items.Item("ChFwYN").Specific.Value =="")
                    {
                        oForm.Items.Item("Btn03").Enabled = false;
                    }

                }
                oMat01.AutoResizeColumns();
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
        /// ?âÏ∂îÍ∞Ä
        /// </summary>
        /// <param name="pSeq"></param>
        /// <param name="oRow"></param>
        private void PS_CO800_AddMatrixRow(short pSeq, int oRow)
        {
            try
            {
                switch (pSeq)
                {
                    case 0:
                        oMat01.AddRow();
                        oDS_PS_CO800L.SetValue("U_LineNum", oRow, (oRow + 1).ToString());
                        oMat01.LoadFromDataSource();
                        break;
                    case 1:
                        oDS_PS_CO800L.InsertRecord(oRow);
                        oDS_PS_CO800L.SetValue("U_LineNum", oRow, (oRow + 1).ToString());
                        oMat01.LoadFromDataSource();
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FormClear
        /// </summary>
        private void PS_CO800_FormClear()
        {
            string DocNum;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_CO800'", "");

                if (Convert.ToDouble(DocNum) == 0)
                {
                    oDS_PS_CO800H.SetValue("DocNum", 0, "1");
                }
                else
                {
                    oDS_PS_CO800H.SetValue("DocNum", 0, DocNum);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Matrix ÎßàÏ?Îß? ÎπàÌñâ ??†ú
        /// </summary>
        private void PS_CO800_Delete_EmptyRow()
        {
            try
            {
                oMat01.FlushToDataSource();

                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_CO800L.GetValue("U_ItemCode", i).ToString().Trim()))
                    {
                        oDS_PS_CO800L.RemoveRecord(i);
                    }
                }

                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Header ?ÑÏàò ?ÖÎ†• ?ÑÎìú Ï≤¥ÌÅ¨
        /// </summary>
        /// <returns></returns>
        private bool PS_CO800_HeaderSpaceLineDel()
        {
            bool returnValue = false;
            string errCode = string.Empty;

            try
            {
                if (oDS_PS_CO800H.GetValue("U_BPLId", 0) == "" )
                {
                    errCode = "1";
                    throw new Exception();
                }
                else if(oDS_PS_CO800H.GetValue("U_DocDate", 0) == "")
                {
                    errCode = "2";
                    throw new Exception();
                }
                else if (oDS_PS_CO800H.GetValue("U_MstCode", 0) == "")
                {
                    errCode = "3";
                    throw new Exception();
                }
                else if (oDS_PS_CO800H.GetValue("U_DoNumber", 0) == "")
                {
                    errCode = "4";
                    throw new Exception();
                }

                returnValue = true;
            }
            
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("?¨ÏóÖ?? ?ÑÏàò?ÖÎ†• ?¨Ìï≠?ÖÎãà??. ?ïÏù∏?òÏÑ∏??.");
                }
               else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox("Í∏∞Ï??ºÏ? ?ÑÏàò?ÖÎ†• ?¨Ìï≠?ÖÎãà??. ?ïÏù∏?òÏÑ∏??.");
                }
                else if (errCode == "3")
                {
                    PSH_Globals.SBO_Application.MessageBox("?¥Îãπ?êÎäî ?ÑÏàò?ÖÎ†• ?¨Ìï≠?ÖÎãà??. ?ïÏù∏?òÏÑ∏??.");
                }
                else if (errCode == "4")
                {
                    PSH_Globals.SBO_Application.MessageBox("?ÑÏûêÍ≤∞Ïû¨Î≤àÌò∏?? ?ÑÏàò?ÖÎ†• ?¨Ìï≠?ÖÎãà??. ?ïÏù∏?òÏÑ∏??.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// Line ?ÑÏàò ?ÖÎ†• ?ÑÎìú Ï≤¥ÌÅ¨
        /// </summary>
        /// <returns></returns>
        private bool PS_CO800_MatrixSpaceLineDel()
        {
            bool returnValue = false;

            int i=0;
            string errCode = string.Empty;

            try
            {
                oMat01.FlushToDataSource();

                //?ºÏù∏
                if (oMat01.VisualRowCount < 1)
                {
                    errCode = "1";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount > 0)
                {
                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                    {
                        oDS_PS_CO800L.Offset = i;
                        if (string.IsNullOrEmpty(oDS_PS_CO800L.GetValue("U_BatchNum", i)))
                        {
                            errCode = "2";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PS_CO800L.GetValue("U_RitmCode", i)))
                        {
                            errCode = "3";
                            throw new Exception();
                        }
                        //else if (string.IsNullOrEmpty(oDS_PS_CO800L.GetValue("U_MoveQty", i)))
                        else if ((oDS_PS_CO800L.GetValue("U_MoveQty", i) == "" ? "0" : oDS_PS_CO800L.GetValue("U_MoveQty", i)) == "0")
                        //  (oRevOffice != "" ? oRevOffice.Substring(0, 3) : oRevOffice);
                        {
                            errCode = "4";
                            throw new Exception();
                        }
                        else if (Convert.ToInt32(oDS_PS_CO800L.GetValue("U_Quantity", i)) < Convert.ToInt32(oDS_PS_CO800L.GetValue("U_MoveQty", i)))
                        {
                            errCode = "4";
                            throw new Exception();
                        }

                    }
                }

                if (string.IsNullOrEmpty(oDS_PS_CO800L.GetValue("U_ItemCode", oMat01.VisualRowCount - 1)))
                {
                    oDS_PS_CO800L.RemoveRecord(oMat01.VisualRowCount - 1);
                }

                oMat01.LoadFromDataSource();

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("?ºÏù∏ ?∞Ïù¥?∞Í? ?ÜÏäµ?àÎã§. ?ïÏù∏?òÏÑ∏??.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox( i + 1  + "?âÏóê Î∞∞ÏπòÎ≤àÌò∏?? ?ÑÏàò?ÖÎãà??.");
                    oMat01.Columns.Item("BatchNum").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "3")
                {
                    PSH_Globals.SBO_Application.MessageBox(i + 1 + "?âÏóê  ?êÏû¨Î£åÏΩî?úÎäî ?ÑÏàò?ÖÎãà??.");
                    oMat01.Columns.Item("RitmCode").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "4")
                {
                    PSH_Globals.SBO_Application.MessageBox(i + 1 + "?? ?¥Îèô ??/Ï§ëÎüâ?? ?ïÏù∏?òÏÑ∏??.");
                    oMat01.Columns.Item("MoveQty").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Î∂ÑÍ∞ú DI
        /// </summary>
        /// <param name="ChkType"></param>
        /// <returns></returns>
        private bool PS_CO800_Create_oDIObject(short ChkType)
        {
            bool returnValue = false;

            int i;
            int j;
            int RetVal;
            int errDiCode;
            double sumAmt;
            string errDiMsg;
            string sTransIdFW = string.Empty;
            string sTransIdRV;
            string sQry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Documents oDIObjectFW = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
            SAPbobsCOM.Documents oDIObjectRV = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();

                oMat01.FlushToDataSource();

                j = 0;
                oDIObjectFW.DocDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
                oDIObjectFW.UserFields.Fields.Item("U_Comments").Value = "Convert Meterial";

                for (i = 1; i < oMat01.VisualRowCount; i++)
                {
                    oDIObjectFW.Lines.Add();
                    oDIObjectFW.Lines.SetCurrentLine(j);
                    oDIObjectFW.Lines.ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim();
                    oDIObjectFW.Lines.Quantity = float.Parse(oMat01.Columns.Item("MoveQty").Cells.Item(i).Specific.Value);
                    oDIObjectFW.Lines.WarehouseCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value.ToString().Trim();
                    oDIObjectFW.Lines.Price = float.Parse(oMat01.Columns.Item("SPrice").Cells.Item(i).Specific.Value);
                    oDIObjectFW.Lines.UnitPrice = float.Parse(oMat01.Columns.Item("SPrice").Cells.Item(i).Specific.Value);
                    oDIObjectFW.Lines.LineTotal = float.Parse(oMat01.Columns.Item("SPrice").Cells.Item(i).Specific.Value) * float.Parse(oMat01.Columns.Item("MoveQty").Cells.Item(i).Specific.Value);
                    oDIObjectFW.Lines.BatchNumbers.BatchNumber = oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value;
                    oDIObjectFW.Lines.BatchNumbers.Quantity = float.Parse(oMat01.Columns.Item("MoveQty").Cells.Item(i).Specific.Value);
                    j += 1;
                }
                
                RetVal = oDIObjectFW.Add();
                if (RetVal != 0)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errMessage = "DI?§Ìñâ Ï§? ?§Î•ò Î∞úÏÉù [" + errDiCode + "][" + errDiMsg + "]";
                    throw new Exception();
                }

                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out sTransIdFW);
                }

                j = 0;
                oDIObjectRV.DocDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
                oDIObjectRV.UserFields.Fields.Item("U_Comments").Value = "Convert Meterial";

                for (i = 1; i < oMat01.VisualRowCount; i++)
                {
                    oDIObjectRV.Lines.Add();
                    oDIObjectRV.Lines.SetCurrentLine(j);
                    oDIObjectRV.Lines.ItemCode = oMat01.Columns.Item("RitmCode").Cells.Item(i).Specific.Value.ToString().Trim();
                    oDIObjectRV.Lines.Quantity = float.Parse(oMat01.Columns.Item("MoveQty").Cells.Item(i).Specific.Value);
                    oDIObjectRV.Lines.WarehouseCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value.ToString().Trim();
                    oDIObjectRV.Lines.Price = float.Parse(oMat01.Columns.Item("SPrice").Cells.Item(i).Specific.Value);
                    oDIObjectRV.Lines.UnitPrice = float.Parse(oMat01.Columns.Item("SPrice").Cells.Item(i).Specific.Value);
                    oDIObjectRV.Lines.LineTotal = float.Parse(oMat01.Columns.Item("SPrice").Cells.Item(i).Specific.Value) * float.Parse(oMat01.Columns.Item("MoveQty").Cells.Item(i).Specific.Value);
                    j += 1;
                }

                RetVal = oDIObjectRV.Add();
                if (RetVal != 0)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errMessage = "DI?§Ìñâ Ï§? ?§Î•ò Î∞úÏÉù [" + errDiCode + "][" + errDiMsg + "]";
                    throw new Exception();
                }

                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    PSH_Globals.oCompany.GetNewObjectCode(out sTransIdRV);

                    sQry = "select sum((a.U_Price - a.U_SPrice) * a.U_MoveQty) as amt ";
                    sQry += " from [@PS_CO800L] a inner join [@PS_CO800H] b on a.DocEntry = b.DocEntry ";
                    sQry += "where b.DocNum = '" + oDS_PS_CO800H.GetValue("DocNum", 0).ToString().Trim() + "'";
                    oRecordSet01.DoQuery(sQry);

                    sumAmt = oRecordSet01.Fields.Item(0).Value;

                    sQry = "  Update [@PS_CO800H] Set U_ChFwYN = 'Y', U_OIGNFw = '" + sTransIdRV + "', U_OIGEFw = '" + sTransIdFW + "', U_Amount = " + sumAmt;
                    sQry += " Where DocNum = '" + oDS_PS_CO800H.GetValue("DocNum", 0).ToString().Trim() + "'";
                    oRecordSet01.DoQuery(sQry);

                    oDS_PS_CO800H.SetValue("U_OIGNFw", 0, sTransIdRV);
                    oDS_PS_CO800H.SetValue("U_OIGEFw", 0, sTransIdFW);
                    oDS_PS_CO800H.SetValue("U_Amount", 0, sumAmt.ToString().Trim());
                    oDS_PS_CO800H.SetValue("U_ChFwYN", 0, "Y");
                }
                oForm.Items.Item("Btn02").Enabled = false;
                oForm.Items.Item("Btn03").Enabled = true;

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObjectFW);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObjectRV);
            }

            return returnValue;
        }

        /// <summary>
        /// Î∂ÑÍ∞ú DI
        /// </summary>
        /// <param name="ChkType"></param>
        /// <returns></returns>
        private bool PS_CO800_Cancel_oDIObject(short ChkType)
        {
            bool returnValue = false;

            int i;
            int j;
            int RetVal;
            int errCode = 0;
            int errDiCode = 0;
            string errDiMsg = string.Empty;
            string ErrLine = string.Empty;
            string sQry;
            string sTransIdFW;
            string sTransIdRV = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Documents oDIObjectFW = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit); //Ï∂úÍ≥†
            SAPbobsCOM.Documents oDIObjectRV = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry); //?ÖÍ≥†

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();

                oMat01.FlushToDataSource();

                j = 0;
                oDIObjectRV.DocDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
                oDIObjectRV.UserFields.Fields.Item("U_CancDoc").Value = oForm.Items.Item("OIGEFw").Specific.Value;
                oDIObjectRV.UserFields.Fields.Item("U_Comments").Value = "Convert Meterial Reverse";
                for (i = 1; i < oMat01.VisualRowCount; i++)
                {
                    oDIObjectRV.Lines.Add();
                    oDIObjectRV.Lines.SetCurrentLine(j);
                    oDIObjectRV.Lines.ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim();
                    oDIObjectRV.Lines.Quantity = float.Parse(oMat01.Columns.Item("MoveQty").Cells.Item(i).Specific.Value);
                    oDIObjectRV.Lines.WarehouseCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value.ToString().Trim();
                    oDIObjectRV.Lines.Price = float.Parse(oMat01.Columns.Item("SPrice").Cells.Item(i).Specific.Value);
                    oDIObjectRV.Lines.UnitPrice = float.Parse(oMat01.Columns.Item("SPrice").Cells.Item(i).Specific.Value);
                    oDIObjectRV.Lines.LineTotal = float.Parse(oMat01.Columns.Item("SPrice").Cells.Item(i).Specific.Value) * float.Parse(oMat01.Columns.Item("MoveQty").Cells.Item(i).Specific.Value);
                    oDIObjectRV.Lines.BatchNumbers.BatchNumber = oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value;
                    oDIObjectRV.Lines.BatchNumbers.Quantity = float.Parse(oMat01.Columns.Item("MoveQty").Cells.Item(i).Specific.Value);

                    j += 1;
                }

                RetVal = oDIObjectRV.Add();
                if (RetVal != 0)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = 1;
                    throw new Exception();
                }

                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out sTransIdRV);
                }

                j = 0;
                oDIObjectFW.DocDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
                oDIObjectFW.UserFields.Fields.Item("U_CancDoc").Value = oForm.Items.Item("OIGNFw").Specific.Value;
                oDIObjectFW.UserFields.Fields.Item("U_Comments").Value = "Convert Meterial Reverse";
                for (i = 1; i < oMat01.VisualRowCount; i++)
                {
                    oDIObjectFW.Lines.Add();
                    oDIObjectFW.Lines.SetCurrentLine(j);
                    oDIObjectFW.Lines.ItemCode = oMat01.Columns.Item("RitmCode").Cells.Item(i).Specific.Value.ToString().Trim();
                    oDIObjectFW.Lines.Quantity = float.Parse(oMat01.Columns.Item("MoveQty").Cells.Item(i).Specific.Value);
                    oDIObjectFW.Lines.WarehouseCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value.ToString().Trim();
                    oDIObjectFW.Lines.Price = float.Parse(oMat01.Columns.Item("SPrice").Cells.Item(i).Specific.Value);
                    oDIObjectFW.Lines.UnitPrice = float.Parse(oMat01.Columns.Item("SPrice").Cells.Item(i).Specific.Value);
                    oDIObjectFW.Lines.LineTotal = float.Parse(oMat01.Columns.Item("SPrice").Cells.Item(i).Specific.Value) * float.Parse(oMat01.Columns.Item("MoveQty").Cells.Item(i).Specific.Value);

                    j += 1;
                }

                //?ÑÎ£å
                RetVal = oDIObjectFW.Add();
                if (RetVal != 0)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = 1;
                    throw new Exception();
                }

                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    PSH_Globals.oCompany.GetNewObjectCode(out sTransIdFW);

                    sQry = "  Update [@PS_CO800H] Set U_ChRvYN = 'Y', U_OIGNRv = '" + sTransIdFW + "', U_OIGERv = '" +  sTransIdRV + "'";
                    sQry += " Where DocNum = '" + oDS_PS_CO800H.GetValue("DocNum", 0).ToString().Trim() + "'";
                    oRecordSet01.DoQuery(sQry);

                    oDS_PS_CO800H.SetValue("U_OIGNRv", 0, sTransIdFW);
                    oDS_PS_CO800H.SetValue("U_OIGERv", 0, sTransIdRV);
                    oDS_PS_CO800H.SetValue("U_ChRvYN", 0, "Y");
                }

                oForm.Items.Item("Btn03").Enabled = false;

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errCode == 7)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("?ÅÏöî?? 50Í∏Ä?? Ï¥àÍ≥º ?±Î°ù Î∂àÍ??©Îãà??. (" + ErrLine + "Î≤àÏß∏ ?ºÏù∏)", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("DI?§Ìñâ Ï§? ?§Î•ò Î∞úÏÉù : [" + errDiCode + "]" + errDiMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObjectFW);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObjectRV);
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

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                    //Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                    //Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    break;
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

                            if (PS_CO800_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_CO800_MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        PS_CO800_Delete_EmptyRow();
                    }
                    else if (pVal.ItemUID == "Btn02")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (PS_CO800_Create_oDIObject(1) == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else
                        {
                            PSH_Globals.SBO_Application.MessageBox("?ïÏù∏ Î™®Îìú?êÏÑú Î≥Ä?òÏù¥ Ï≤òÎ¶¨?òÏÑ∏??.");
                        }
                    }
                    else if (pVal.ItemUID == "Btn03")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && oForm.Items.Item("ChFwYN").Specific.Value == "Y")
                        {
                            if (PS_CO800_Cancel_oDIObject(1) == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else
                        {
                            PSH_Globals.SBO_Application.MessageBox("Î≥Ä?? Ï∑®ÏÜåÍ∞Ä ?†Ìñâ?? ?? ?ÜÏäµ?àÎã§.");
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        PS_CO800_AddMatrixRow(1, oMat01.RowCount);
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
        /// KEY_DOWN ?¥Î≤§??
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent Í∞ùÏ≤¥</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ColUID == "ItemCode")
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ColUID == "BatchNum")
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("BatchNum").Cells.Item(pVal.Row).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ColUID == "RitmCode")
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("RitmCode").Cells.Item(pVal.Row).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.Action_Success == true)
                    {
                        oSeq = 1;
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
        /// GOT_FOCUS ?¥Î≤§??
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent Í∞ùÏ≤¥</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    oLastItemUID01 = pVal.ItemUID;
                }
                else if (pVal.Before_Action == false)
                {
                    oLastItemUID01 = pVal.ItemUID;
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
        /// CLICK ?¥Î≤§??
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent Í∞ùÏ≤¥</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
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
        /// DOUBLE_CLICK ?¥Î≤§??
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent Í∞ùÏ≤¥</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        oMat01.FlushToDataSource();

                        if (pVal.ColUID == "Check")
                        {
                            string checkYN;

                            if (oDS_PS_CO800L.GetValue("U_Check", 0).ToString().Trim() == "" || oDS_PS_CO800L.GetValue("U_Check", 0).ToString().Trim() == "N")
                            {
                                checkYN = "Y";
                            }
                            else
                            {
                                checkYN = "N";
                            }

                            for (int i = 0; i <= oDS_PS_CO800L.Size - 1; i++)
                            {
                                oDS_PS_CO800L.SetValue("U_Check", i, checkYN);
                            }
                        }

                        oMat01.LoadFromDataSource();
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
                oForm.Freeze(false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO800H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO800L);
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
        /// FORM_ACTIVATE ?¥Î≤§??
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent Í∞ùÏ≤¥</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_ACTIVATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (oSeq == 1)
                    {
                        oSeq = 0;
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
        /// VALIDATE ?¥Î≤§??
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent Í∞ùÏ≤¥</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            string sQry;
            string price;
            string amount;
            string whsCode;
            string quantity;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
                        switch (pVal.ItemUID)
                        {
                            case "MstCode":
                                oDS_PS_CO800H.SetValue("U_MstName", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", ""));
                                break;

                            case "DocDate":
                                oMat01.FlushToDataSource();
                                for(i = 0; i < oMat01.VisualRowCount-1; i++)
                                {
                                    oDS_PS_CO800L.SetValue("U_BatchNum",i, "");
                                    oDS_PS_CO800L.SetValue("U_WhsCode", i, "");
                                    oDS_PS_CO800L.SetValue("U_Price", i, "");
                                    oDS_PS_CO800L.SetValue("U_Amount", i, "");
                                    oDS_PS_CO800L.SetValue("U_Quantity", i, "");
                                    oDS_PS_CO800L.SetValue("U_RitmCode", i, "");
                                    oDS_PS_CO800L.SetValue("U_RitmName", i, "");
                                    oDS_PS_CO800L.SetValue("U_MoveQty", i, "");
                                }

                                oMat01.LoadFromDataSource();
                                break;

                            case "Mat01":

                                if (pVal.ColUID == "ItemCode")
                                {
                                    sQry =  "SELECT count(*)";
                                    sQry += "  FROM OINV A INNER JOIN INV1 B ON A.DOCENTRY = B.DOCENTRY";
                                    sQry += "  	    	  INNER JOIN DLN1 C ON B.BaseEntry = C.DocEntry and B.BaseLine = C.LineNum";
                                    sQry += "			  INNER JOIN OIVL D ON D.TransType = '15' and C.DocEntry = D.BASE_REF AND C.LineNum = D.DocLineNum";
                                    sQry += " WHERE Convert(varchar(6), a.DocDate,112) = Convert(varchar(6),'" + oForm.Items.Item("DocDate").Specific.Value + "',112)";
                                    sQry += "   AND b.itemcode ='" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value  + "'";
                                    oRecordSet01.DoQuery(sQry);

                                    if (oRecordSet01.Fields.Item(0).Value == 0)
                                    {
                                        errMessage = "?†ÌÉù?òÏã† ?îÏóê ?¥Îãπ ?àÎ™©?? ?©Ìíà?¥Î†•?? ?ÜÏñ¥?? Ï≤òÎ¶¨?? ?? ?ÜÏäµ?àÎã§.(Í∞ÄÍ≥µÎπÑÏ≤òÎ¶¨Î∂àÍ?)";
                                        oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value = "";
                                        throw new Exception();
                                    }
                                    oMat01.FlushToDataSource();

                                    oDS_PS_CO800L.SetValue("U_ItemName", pVal.Row - 1, dataHelpClass.Get_ReData("ItemName", "ItemCode", "OITM", "'" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "'", ""));
                                    oMat01.LoadFromDataSource();
                                }
                                else if (pVal.ColUID == "RitmCode")
                                {
                                    oMat01.FlushToDataSource();

                                    oDS_PS_CO800L.SetValue("U_RitmName", pVal.Row - 1, dataHelpClass.Get_ReData("ItemName", "ItemCode", "OITM", "'" + oMat01.Columns.Item("RitmCode").Cells.Item(pVal.Row).Specific.Value + "'", ""));
                                    oMat01.LoadFromDataSource();
                                }
                                else if (pVal.ColUID == "BatchNum")
                                {
                                    oMat01.FlushToDataSource();

                                    //Î∞∞ÏπòÎ≤àÌò∏ Ï∞ΩÍ≥† Î∞? ?òÎüâ
                                    sQry = "SELECT	  B.WhsCode";
                                    sQry += "		, B.QUANTITY";
                                    sQry += "  FROM OBTN A LEFT JOIN OBTQ B ON A.ABSENTRY = B.MDABSENTRY";
                                    sQry += " WHERE 1=1";
                                    sQry += "  AND B.Quantity <> '0'";
                                    sQry += "  AND a.DISTNUMBER ='" + oMat01.Columns.Item("BatchNum").Cells.Item(pVal.Row).Specific.Value +"'";

                                    oRecordSet01.DoQuery(sQry);

                                    whsCode = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                                    quantity = oRecordSet01.Fields.Item(1).Value.ToString().Trim();

                                    sQry = "select AVG(U_Price)";
                                    sQry += " from [@PS_MM001H] a inner join [@PS_MM001L] b on a.DocEntry = b.DocEntry and a.Canceled ='N' ";
                                    sQry += "where  b.U_ItemCode ='" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "'";
                                    sQry += "  and Convert(varchar(6),a.U_DocDate,112) = Convert(varchar(6),'" + oForm.Items.Item("DocDate").Specific.Value + "',112)";
                                    oRecordSet01.DoQuery(sQry);

                                    oDS_PS_CO800L.SetValue("U_SPrice", pVal.Row - 1, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                                    oDS_PS_CO800L.SetValue("U_SAmount", pVal.Row - 1, (oRecordSet01.Fields.Item(0).Value * Int32.Parse(quantity)).ToString());

                                    // ?®Í? Ï°∞Ìöå
                                    sQry = " SELECT ROUND(SUM(transValue)/(SUM(InQty)-SUM(OutQty)),0) AS PRICE";
                                    sQry += " FROM OINM";
                                    sQry += "  WHERE ITEMCODE ='" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "'";
                                    sQry += "AND DocDate <= DATEADD(d,-1,CONVERT(DATETIME,CONVERT(CHAR(6),DATEADD(m,0,'" + oForm.Items.Item("DocDate").Specific.Value + "'),112) + '01'))";

                                    oRecordSet01.DoQuery(sQry);
                                    price = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                                    amount = (oRecordSet01.Fields.Item(0).Value * Int32.Parse(quantity)).ToString();

                                    oDS_PS_CO800L.SetValue("U_WhsCode", pVal.Row - 1, whsCode);
                                    oDS_PS_CO800L.SetValue("U_Quantity", pVal.Row - 1, quantity);
                                    oDS_PS_CO800L.SetValue("U_Price", pVal.Row - 1, price);
                                    oDS_PS_CO800L.SetValue("U_Amount", pVal.Row - 1, amount);
                                    oMat01.LoadFromDataSource();

                                    if (oMat01.RowCount <= pVal.Row)
                                    {
                                        PS_CO800_AddMatrixRow(1, oMat01.RowCount);
                                    }

                                }
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oMat01.AutoResizeColumns();
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if(errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                BubbleEvent = false;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// ROW_DELETE ?¥Î≤§??
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (int i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }

                        oMat01.FlushToDataSource();
                        oDS_PS_CO800L.RemoveRecord(oDS_PS_CO800L.Size - 1);
                        oMat01.LoadFromDataSource();

                        if (oMat01.RowCount == 0)
                        {
                            PS_CO800_AddMatrixRow(0, 0);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_CO800L.GetValue("U_ItemCode", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_CO800_AddMatrixRow(0, oMat01.RowCount);
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
                        case "1284": //Ï∑®ÏÜå
                            break;
                        case "1286": //?´Í∏∞
                            break;
                        case "1293": //?âÏÇ≠??
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;

                        case "1281": //Ï∞æÍ∏∞
                            break;
                        case "1282": //Ï∂îÍ?
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //?àÏΩî?úÏù¥?ôÎ≤Ñ??
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //Ï∑®ÏÜå
                            break;
                        case "1286": //?´Í∏∞
                            break;
                        case "1281": //Ï∞æÍ∏∞
                            PS_CO800_FormItemEnabled();
                            PS_CO800_AddMatrixRow(1, oMat01.RowCount);

                            break;
                        case "1282": //Ï∂îÍ?
                            PS_CO800_FormItemEnabled();
                            PS_CO800_FormClear();
                            PS_CO800_AddMatrixRow(0, oMat01.RowCount);
                            oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //?àÏΩî?úÏù¥?ôÎ≤Ñ??
                            PS_CO800_FormItemEnabled();
                            PS_CO800_AddMatrixRow(1, oMat01.RowCount);

                            break;
                        case "1293": //?âÏÇ≠??
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            PS_CO800_AddMatrixRow(0, 0);
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
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
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
