﻿using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 구매청구정정처리
    /// </summary>
    internal class PS_MM018 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat;
        private SAPbouiCOM.DBDataSource oDS_PS_MM018H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_MM018L; //등록라인

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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM018.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM018_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM018");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocNum";

                oForm.Freeze(true);
                PS_MM018_CreateItems();
                PS_MM018_ComboBox_Setting();
                PS_MM018_Initialization();
                PS_MM018_FormItemEnabled();
                PS_MM018_FormClear();
                PS_MM018_AddMatrixRow(0, true);

                oForm.EnableMenu("1283", false); // 삭제
                oForm.EnableMenu("1286", false); // 닫기
                oForm.EnableMenu("1287", false); // 복제
                oForm.EnableMenu("1284", true); // 취소
                oForm.EnableMenu("1293", true); // 행삭제
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
        private void PS_MM018_CreateItems()
        {
            try
            {
                oDS_PS_MM018H = oForm.DataSources.DBDataSources.Item("@PS_MM018H");
                oDS_PS_MM018L = oForm.DataSources.DBDataSources.Item("@PS_MM018L");
                oMat = oForm.Items.Item("Mat01").Specific;
                oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat.AutoResizeColumns();

                oDS_PS_MM018H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_MM018_ComboBox_Setting()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("BPLID").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }

                // 승인여부
                oForm.Items.Item("ConfYN").Specific.ValidValues.Add("N", "승인");
                oForm.Items.Item("ConfYN").Specific.ValidValues.Add("Y", "미승인");
                oDS_PS_MM018H.SetValue("U_ConfYN", 0, "Y");

                // 반영여부
                oForm.Items.Item("ApplyYN").Specific.ValidValues.Add("N", "반영");
                oForm.Items.Item("ApplyYN").Specific.ValidValues.Add("Y", "미반영");
                oDS_PS_MM018H.SetValue("U_ApplyYN", 0, "Y");

                // 처리여부
                oMat.Columns.Item("ComplYN").ValidValues.Add("Y", "처리");
                oMat.Columns.Item("ComplYN").ValidValues.Add("N", "미처리");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// Initialization
        /// </summary>
        private void PS_MM018_Initialization()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                /*등록자*/
                sQry = "select USER_CODE  from [OUSR] where userid ='" + PSH_Globals.oCompany.UserSignature + "'";
                oRecordSet.DoQuery(sQry);
                oForm.Items.Item("UserID").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();

                sQry = "select U_USERName from [@PS_SY005L] where code ='MM018' and U_USERID  = '" + oForm.Items.Item("UserID").Specific.Value.ToString().Trim() + "'";
                oRecordSet.DoQuery(sQry);
                oForm.Items.Item("UserName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                
                /*승인자*/
                sQry = "select U_AppUser  from [@PS_SY005L] where code ='MM018' and U_USERID  = '" + oForm.Items.Item("UserID").Specific.Value.ToString().Trim() + "'";
                oRecordSet.DoQuery(sQry);
                oForm.Items.Item("ConfID").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();

                sQry = "select U_AppName from [@PS_SY005L] where code ='MM018' and U_AppUser = '" + oForm.Items.Item("ConfID").Specific.Value.ToString().Trim() + "'";
                oRecordSet.DoQuery(sQry);
                oForm.Items.Item("ConfName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// HeaderSpaceLineDeldev
        /// </summary>
        /// <returns></returns>
        private bool PS_MM018_HeaderSpaceLineDel()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_MM018H.GetValue("U_BPLID", 0)))
                {
                    errMessage = "사업장은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM018H.GetValue("U_UserID", 0)))
                {
                    errMessage = "등록자는 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM018H.GetValue("U_DocDate", 0)))
                {
                    errMessage = "등록일자는 필수사항입니다. 확인하세요.";
                    throw new Exception();
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// MatrixSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_MM018_MatrixSpaceLineDel()
        {
            bool returnValue = false;
            int i;
            string errMessage = string.Empty;

            try
            {
                oMat.FlushToDataSource();
                if (oMat.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }
                else if (oMat.VisualRowCount == 1 && string.IsNullOrEmpty(oDS_PS_MM018L.GetValue("U_MM005Doc", 0)))
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                for (i = 0; i <= oMat.VisualRowCount - 2; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_MM018L.GetValue("U_MM005Doc", i)))
                    {
                        errMessage = "구매요청문서번호는 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                }
                oDS_PS_MM018L.RemoveRecord(oDS_PS_MM018L.Size - 1);
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
            return returnValue;
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_MM018_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oForm.Items.Item("BPLID").Enabled = true;
                    oForm.Items.Item("RqstID").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("ApplyBt").Enabled = false;
                    oDS_PS_MM018H.SetValue("U_ConfYN", 0, "Y"); //승인여부 
                    oDS_PS_MM018H.SetValue("U_ApplyYN", 0, "Y"); //승인여부 
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;
                    oForm.Items.Item("BPLID").Enabled = false;
                    oForm.Items.Item("RqstID").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oForm.Items.Item("ApplyBt").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("BPLID").Enabled = false;
                    oForm.Items.Item("RqstID").Enabled = true;
                    oForm.Items.Item("Note").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = false;

                    if(oForm.Items.Item("ConfYN").Specific.Value == "N")
                    {
                        oForm.Items.Item("Mat01").Enabled = false;
                        oForm.Items.Item("ApplyBt").Enabled = true;
                        oForm.Items.Item("RqstID").Enabled = false;
                        oForm.Items.Item("Note").Enabled = false;

                        if (oForm.Items.Item("ApplyYN").Specific.Value == "N")
                        {
                            oForm.Items.Item("ApplyBt").Enabled = false;
                        }
                    }
                    else
                    {
                        oForm.Items.Item("Mat01").Enabled = true;
                        oForm.Items.Item("ApplyBt").Enabled = false;
                        oForm.Items.Item("RqstID").Enabled = true;
                        oForm.Items.Item("Note").Enabled = true;
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
        /// PS_MM018_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_MM018_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_MM018L.InsertRecord(oRow);
                }
                oMat.AddRow();
                oDS_PS_MM018L.Offset = oRow;
                oDS_PS_MM018L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat.LoadFromDataSource();
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
        private void PS_MM018_FormClear()
        {
            string DocNum;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM018'", "");
                if (Convert.ToDouble(DocNum) == 0)
                {
                    oForm.Items.Item("DocNum").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocNum").Specific.Value = DocNum;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_MM018_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                switch (oUID)
                {
                    case "RqstID":
                        sQry = "Select U_FullName From [@PH_PY001A] Where Code = '" + oForm.Items.Item("RqstID").Specific.Value.ToString().Trim() + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Items.Item("RqstName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                        break;

                    case "Mat01":
                        if (oCol == "MM005Doc")
                        {
                            oForm.Freeze(true);
                            oMat.FlushToDataSource();
                            sQry = "exec [PS_MM018_01] '" + oMat.Columns.Item("MM005Doc").Cells.Item(oRow).Specific.Value + "'";
                            oRecordSet.DoQuery(sQry);

                            if (oRecordSet.RecordCount > 0)
                            {
                                oDS_PS_MM018L.SetValue("U_MM005Qty", oRow -1, oRecordSet.Fields.Item(1).Value);
                                oDS_PS_MM018L.SetValue("U_MM010Doc", oRow - 1, oRecordSet.Fields.Item(2).Value);
                                oDS_PS_MM018L.SetValue("U_MM010Lin", oRow - 1, oRecordSet.Fields.Item(3).Value);
                                oDS_PS_MM018L.SetValue("U_MM010Qty", oRow - 1, oRecordSet.Fields.Item(4).Value);
                                oDS_PS_MM018L.SetValue("U_OPOROrd", oRow - 1, oRecordSet.Fields.Item(5).Value);
                                oDS_PS_MM018L.SetValue("U_OPORLine", oRow - 1, oRecordSet.Fields.Item(6).Value);
                                oDS_PS_MM018L.SetValue("U_OPORQty", oRow - 1, oRecordSet.Fields.Item(7).Value);
                                oDS_PS_MM018L.SetValue("U_MM030Doc", oRow - 1, oRecordSet.Fields.Item(8).Value);
                                oDS_PS_MM018L.SetValue("U_MM030Lin", oRow - 1, oRecordSet.Fields.Item(9).Value);
                                oDS_PS_MM018L.SetValue("U_MM030Qty", oRow - 1, oRecordSet.Fields.Item(10).Value);
                            }
                            oMat.LoadFromDataSource();
                            oMat.AutoResizeColumns();
                            oForm.Freeze(false);
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_MM018_PurchaseOrders_Close
        /// </summary>
        /// <returns></returns>
        private bool PS_MM018_PurchaseOrders_Close()
        {
            bool returnValue = false;
            int i;
            int RetVal;
            int errDICode;
            int OPORLine;
            string sQry;
            string errDIMsg;
            string OPOROrd;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Documents DI_oPurchaseOrders = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

            try
            {
                for(i= 0; i < oMat.RowCount -1; i++)
                {
                    OPOROrd = oDS_PS_MM018L.GetValue("U_OPOROrd", i).ToString().Trim();
                    OPORLine = Convert.ToInt32(oDS_PS_MM018L.GetValue("U_OPORLine", i).ToString().Trim());
                    if (DI_oPurchaseOrders.GetByKey(Convert.ToInt32(OPOROrd)) == false) //완료
                    {
                        PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                        errMessage = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                        throw new Exception();
                    }
                    DI_oPurchaseOrders.Lines.SetCurrentLine(OPORLine);
                    DI_oPurchaseOrders.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close;
                    RetVal = DI_oPurchaseOrders.Update();
                    if (0 != RetVal)
                    {
                        PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                        errMessage = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                        throw new Exception();
                    }
                    else
                    {
                        sQry = "UPDATE [@PS_MM018L] SET U_ComplYN = 'Y' WHERE DocEntry ='" + oForm.Items.Item("DocNum").Specific.Value.ToString().Trim() + "' and U_MM005Doc ='" + oDS_PS_MM018L.GetValue("U_MM005Doc", i).ToString().Trim() + "' ";
                        oRecordSet.DoQuery(sQry);
                    }
                }
                sQry = "UPDATE [@PS_MM018H] SET U_ApplyYN = 'N' WHERE DocEntry ='" + oForm.Items.Item("DocNum").Specific.Value.ToString().Trim() + "'";
                oRecordSet.DoQuery(sQry);
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oPurchaseOrders);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                            if (PS_MM018_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_MM018_MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        oMat.AutoResizeColumns();
                    }
                    else if(pVal.ItemUID  == "ApplyBt")
                    {
                        PS_MM018_PurchaseOrders_Close();
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
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "UserID")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("UserID").Specific.Value.ToString().Trim()))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "MM005Doc")
                            {
                                if (string.IsNullOrEmpty(oMat.Columns.Item("MM005Doc").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "RqstID")
                        {
                            PS_MM018_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "MM005Doc")
                            {
                                PS_MM018_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                                if (oMat.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_MM018L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_MM018_AddMatrixRow(pVal.Row, false);
                                }
                                oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                BubbleEvent = false;
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
            int i;

            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat.VisualRowCount; i++)
                        {
                            oMat.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        oMat.FlushToDataSource();
                        oDS_PS_MM018L.RemoveRecord(oDS_PS_MM018L.Size - 1);
                        oMat.LoadFromDataSource();
                        if (oMat.RowCount == 0)
                        {
                            PS_MM018_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_MM018L.GetValue("U_MM005Doc", oMat.RowCount - 1).ToString().Trim()))
                            {
                                PS_MM018_AddMatrixRow(oMat.RowCount, false);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    SubMain.Remove_Forms(oFormUniqueID);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM018H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM018L);
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
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_MM018_FormItemEnabled();
                            oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                            break;
                        case "1282": //추가
                            PS_MM018_Initialization();
                            PS_MM018_FormItemEnabled();
                            PS_MM018_FormClear();
                            oDS_PS_MM018H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                            PS_MM018_AddMatrixRow(0, true);
                            oMat.AutoResizeColumns();
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                            PS_MM018_FormItemEnabled();
                            if (oMat.VisualRowCount > 0)
                            {
                                if (!string.IsNullOrEmpty(oMat.Columns.Item("MM005Doc").Cells.Item(oMat.VisualRowCount).Specific.Value))
                                {
                                    if (oDS_PS_MM018H.GetValue("Status", 0) == "O")
                                    {
                                        PS_MM018_AddMatrixRow(oMat.RowCount, false);
                                    }
                                }
                            }
                            oMat.AutoResizeColumns();
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
