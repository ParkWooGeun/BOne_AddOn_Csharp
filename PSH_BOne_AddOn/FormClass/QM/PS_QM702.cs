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
    /// 검사승인 및 전송
    /// </summary>
    internal class PS_QM702 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_QM702H; //등록헤더
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.DBDataSource oDS_PS_QM702M;

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM702.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_QM702_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_QM702");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

                oForm.Freeze(true);
                PS_QM702_CreateItems();
                PS_QM702_AddMatrixRowM(0, true);
                PS_QM702_AddMatrixRow(0,true);
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_QM702_CreateItems()
        {
            try
            {
                oDS_PS_QM702H = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oDS_PS_QM702M = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");

                // 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("oMat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                // 메트릭스 개체 할당
                oMat02 = oForm.Items.Item("oMat02").Specific;
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat02.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }
      
        /// <summary>
        /// 매트릭스 행 추가
        /// PH_PY035_Add_MatrixRow
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        /// </summary>
        private void PS_QM702_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_QM702H.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_QM702H.Offset = oRow;
                oDS_PS_QM702H.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PS_QM702H_Add_MatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
        /// <summary>
        /// 매트릭스 행 추가
        /// PH_PY035_Add_MatrixRow
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        /// </summary>
        private void PS_QM702_AddMatrixRowM(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_QM702M.InsertRecord(oRow);
                }
                oMat02.AddRow();
                oDS_PS_QM702M.Offset = oRow;
                oDS_PS_QM702M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat02.LoadFromDataSource();
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PS_QM702M_Add_MatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
        /// <summary>
        /// LoadData
        /// </summary>
        private void PS_QM702_LoadData(string Gubun)
        {
            string sQry;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();
                oDS_PS_QM702H.Clear(); //추가

                sQry = "EXEC [PS_QM702_01] '" + Gubun + "'";
                oRecordSet01.DoQuery(sQry);

                if(Gubun == "O")
                {
                    for (int i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                    {
                        if (i + 1 > oDS_PS_QM702H.Size)
                        {
                            oDS_PS_QM702H.InsertRecord((i));
                        }
                        oMat01.AddRow();
                        oDS_PS_QM702H.Offset = i;
                        oDS_PS_QM702H.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                        oDS_PS_QM702H.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("U_InOut").Value.ToString().Trim());
                        oDS_PS_QM702H.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());
                        oDS_PS_QM702H.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("U_CLTCOD").Value.ToString().Trim());
                        oDS_PS_QM702H.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("U_WorkNum").Value.ToString().Trim());
                        oDS_PS_QM702H.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("U_ItemName").Value.ToString().Trim());
                        oDS_PS_QM702H.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("U_WorkDate").Value.ToString().Trim());
                        oDS_PS_QM702H.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("U_WorkCode").Value.ToString().Trim());
                        oDS_PS_QM702H.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("U_WorkName").Value.ToString().Trim());
                        oDS_PS_QM702H.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("U_CardCode").Value.ToString().Trim());
                        oDS_PS_QM702H.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("U_CardName").Value.ToString().Trim());
                        oDS_PS_QM702H.SetValue("U_ColReg11", i, oRecordSet01.Fields.Item("U_BZZadQty").Value.ToString().Trim());
                        oDS_PS_QM702H.SetValue("U_ColReg12", i, oRecordSet01.Fields.Item("U_BadCode").Value.ToString().Trim());
                        oDS_PS_QM702H.SetValue("U_ColReg13", i, oRecordSet01.Fields.Item("U_BadNote").Value.ToString().Trim());
                        oDS_PS_QM702H.SetValue("U_ColReg14", i, oRecordSet01.Fields.Item("U_verdict").Value.ToString().Trim());
                        oDS_PS_QM702H.SetValue("U_ColReg15", i, oRecordSet01.Fields.Item("U_Comments").Value.ToString().Trim());
                        oRecordSet01.MoveNext();
                    }
                }
               else
                {
                    for (int i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                    {
                        if (i + 1 > oDS_PS_QM702H.Size)
                        {
                            oDS_PS_QM702H.InsertRecord((i));
                        }
                        oMat01.AddRow();
                        oDS_PS_QM702M.Offset = i;
                        oDS_PS_QM702M.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                        oDS_PS_QM702M.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("U_InOut").Value.ToString().Trim());
                        oDS_PS_QM702M.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());
                        oDS_PS_QM702M.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("U_CLTCOD").Value.ToString().Trim());
                        oDS_PS_QM702M.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("U_WorkNum").Value.ToString().Trim());
                        oDS_PS_QM702M.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("U_ItemName").Value.ToString().Trim());
                        oDS_PS_QM702M.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("U_WorkDate").Value.ToString().Trim());
                        oDS_PS_QM702M.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("U_WorkCode").Value.ToString().Trim());
                        oDS_PS_QM702M.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("U_WorkName").Value.ToString().Trim());
                        oDS_PS_QM702M.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("U_CardCode").Value.ToString().Trim());
                        oDS_PS_QM702M.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("U_CardName").Value.ToString().Trim());
                        oDS_PS_QM702M.SetValue("U_ColReg11", i, oRecordSet01.Fields.Item("U_BZZadQty").Value.ToString().Trim());
                        oDS_PS_QM702M.SetValue("U_ColReg12", i, oRecordSet01.Fields.Item("U_BadCode").Value.ToString().Trim());
                        oDS_PS_QM702M.SetValue("U_ColReg13", i, oRecordSet01.Fields.Item("U_BadNote").Value.ToString().Trim());
                        oDS_PS_QM702M.SetValue("U_ColReg14", i, oRecordSet01.Fields.Item("U_verdict").Value.ToString().Trim());
                        oDS_PS_QM702M.SetValue("U_ColReg15", i, oRecordSet01.Fields.Item("U_Comments").Value.ToString().Trim());
                        oRecordSet01.MoveNext();
                    }
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (System.Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY702_MTX01:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
        }
        /// <summary>
        /// LoadData
        /// </summary>
        private bool PS_QM702_UPDATEData(string DocEntry, string Gobun)
        {
            string sQry;
            string errMessage = string.Empty;
            bool returnValue = false;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (Gobun == "외주")
                {
                    sQry = "UPDATE [@PS_QM701H] SET U_ChkYN = '승인', U_ChkDate = Convert(CHAR(10),GETDATE()) WHERE DocEntry ='" + DocEntry + "'";
                }
                else
                {
                    sQry = "UPDATE [@PS_QM703H] SET U_ChkYN = '승인', U_ChkDate = Convert(CHAR(10),GETDATE()) WHERE DocEntry ='" + DocEntry + "'";
                }
                
                oRecordSet01.DoQuery(sQry);
                returnValue = true;
            }
            catch (System.Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY035_MTX01:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
            }
            return returnValue;
        }

        /// <summary>
        /// Send_EMail
        /// </summary>
        /// <param name="p_DocEntry"></param>
        /// <param name="p_Reson"></param>
        /// <returns></returns>
        private bool Return_EMail(string p_DocEntry, string p_Email, string p_Reson, string p_gobun)
        {
            bool ReturnValue = false;
            string strToAddress;
            string strSubject;
            string strBody;
            string sQry;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                strSubject = "부적합문서 반려";
                strBody = "부적합  " + p_gobun + "  문서번호" + p_DocEntry + "가 반려되었습니다. ";
                strBody += "반려사유 : " + p_Reson + "입니다.";

                strToAddress = p_Email;

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
                mail.Send();

                mail = null;
                outlookApp = null;

                if (p_gobun == "외주")
                {
                    sQry = "UPDATE [@PS_QM701H] SET U_ChkYN = '반려', U_ChkDate = Convert(CHAR(10),GETDATE()) WHERE DocEntry ='" + p_DocEntry + "'";
                }
                else
                {
                    sQry = "UPDATE [@PS_QM703H] SET U_ChkYN = '반려', U_ChkDate = Convert(CHAR(10),GETDATE()) WHERE DocEntry ='" + p_DocEntry + "'";
                }
                oRecordSet01.DoQuery(sQry);
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

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                    //    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

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
                    PS_QM702_AddMatrixRowM(oMat02.VisualRowCount, false);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "btn_search")
                    {
                        PS_QM702_LoadData("I");
                        PS_QM702_LoadData("O");
                    }
                    else if (pVal.ItemUID == "btn_appr")
                    {
                        oMat01.FlushToDataSource();
                        for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            if (oDS_PS_QM702H.GetValue("U_ColReg17", i).ToString().Trim() == "Y")
                            {
                                string GOBUN = oDS_PS_QM702H.GetValue("U_ColReg01", i).ToString().Trim();
                                string DocEntry = oDS_PS_QM702H.GetValue("U_ColReg02", i).ToString().Trim();
                                if (PS_QM702_UPDATEData(DocEntry,    GOBUN) == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }
                    }
                    else if (pVal.ItemUID == "btn_return")
                    {
                        oMat01.FlushToDataSource();
                        for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            if (oDS_PS_QM702H.GetValue("U_ColReg17", i).ToString().Trim() == "Y")
                            {
                                if (string.IsNullOrEmpty(oDS_PS_QM702H.GetValue("U_ColReg16", i).ToString().Trim()))
                                {
                                    errMessage = "반려 시 반려사유는 필수입니다.";
                                    throw new Exception();
                                }
                                else
                                {
                                    string DocEntry = oDS_PS_QM702H.GetValue("U_ColReg02", i).ToString().Trim();
                                    string Reson = oDS_PS_QM702H.GetValue("U_ColReg16", i).ToString().Trim();
                                    string GOBUN = oDS_PS_QM702H.GetValue("U_ColReg01", i).ToString().Trim();

                                    sQry = "SELECT U_eMail FROM [@PS_QM700L] WHERE U_UseYN = 'Y'AND Code ='ZReturn'";
                                    oRecordSet01.DoQuery(sQry);


                                    for (int j = 0; j <= oRecordSet01.RecordCount - 1; j++)
                                    {
                                        string email = string.Empty;
                                        email = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                                        if (Return_EMail(DocEntry, email, Reson, GOBUN) == false)//사번
                                        {
                                            errMessage = "전송 중 오류가 발생했습니다.";
                                            throw new Exception();
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else if (pVal.BeforeAction == false)
                    {
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
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
            oForm.Freeze(false);
        }

        /// <summary>
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "oMat01")
                    {
                        if (pVal.ColUID == "DocEntry")
                            {
                            PS_QM701 tempForm = new PS_QM701();
                            tempForm.LoadForm(oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value);
                        }
                    }
                    else if (pVal.ItemUID == "oMat02")
                    {
                        if (pVal.ColUID == "DocEntry")
                        {
                            PS_QM703 tempForm = new PS_QM703();
                            tempForm.LoadForm(oMat02.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value);
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
                if (pVal.ItemUID == "Mat01")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = pVal.ColUID;
                        oLastColRow01 = pVal.Row;
                    }
                }
                else
                {
                    oLastItemUID01 = pVal.ItemUID;
                    oLastColUID01 = "";
                    oLastColRow01 = 0;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM702H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM702M);
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
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
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
                            break;
                        case "1281": //찾기
                        case "1282": //추가
                            PS_QM702_AddMatrixRowM(0, true);
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }
    }
}
