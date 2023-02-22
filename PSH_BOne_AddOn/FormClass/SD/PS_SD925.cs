using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 멀티 단가변경처리
    /// </summary>
    internal class PS_SD925 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_SD925H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_SD925L; //등록라인
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD925.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_SD925_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_SD925");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_SD925_CreateItems();
                PS_SD925_ComboBox_Setting();
                PS_SD925_Initial_Setting();
                PS_SD925_FormItemEnabled();
                PS_SD925_FormClear();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>        
        private void PS_SD925_CreateItems()
        {
            try
            {
                oDS_PS_SD925H = oForm.DataSources.DBDataSources.Item("@PS_SD925H");
                oDS_PS_SD925L = oForm.DataSources.DBDataSources.Item("@PS_SD925L");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_SD925_ComboBox_Setting()
        {
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oRecordSet01.DoQuery("SELECT BPLId, BPLName From[OBPL] order by 1"); //사업장
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
        /// 화면 초기화
        /// </summary>
        private void PS_SD925_Initial_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FormClear
        /// </summary>
        private void PS_SD925_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD925'", "");
                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.Value = "1";
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 각 모드에 따른 아이템설정
        /// </summary>
        private void PS_SD925_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("YYYYMM").Enabled = true;
                    oForm.Items.Item("Btn01").Enabled = true;
                    oForm.Items.Item("TgtWgt").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oMat01.Columns.Item("Check").Editable = true;
                    oForm.Items.Item("ChgPriBt").Enabled = false;
                    oForm.Items.Item("ReChgPri").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("YYYYMM").Enabled = true;
                    oForm.Items.Item("TgtWgt").Enabled = false;
                    oForm.Items.Item("Btn01").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oMat01.Columns.Item("Check").Editable = false;
                    oForm.Items.Item("ChgPriBt").Enabled = false;
                    oForm.Items.Item("ReChgPri").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("YYYYMM").Enabled = false;
                    oForm.Items.Item("TgtWgt").Enabled = false;
                    oForm.Items.Item("Btn01").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    if (oForm.Items.Item("U_Status").Specific.Value == "Y")
                    {
                        oMat01.Columns.Item("Check").Editable = false;
                        oForm.Items.Item("ChgPriBt").Enabled = false;
                        oForm.Items.Item("ReChgPri").Enabled = true;
                    }
                    else if(oForm.Items.Item("U_Status").Specific.Value == "C")
                    {
                        oMat01.Columns.Item("Check").Editable = false;
                        oForm.Items.Item("ChgPriBt").Enabled = false;
                        oForm.Items.Item("ReChgPri").Enabled = false;
                    }
                    else
                    {
                        oMat01.Columns.Item("Check").Editable = true;
                        oForm.Items.Item("ChgPriBt").Enabled = true;
                        oForm.Items.Item("ReChgPri").Enabled = false;
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
        /// 메트릭스 Row추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_SD925_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                oMat01.FlushToDataSource();
                if (RowIserted == false)
                {
                    oDS_PS_SD925L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_SD925L.Offset = oRow;
                oDS_PS_SD925L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
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
        /// 필수입력사항 체크(헤더)
        /// </summary>
        /// <returns></returns>
        private bool PS_SD925_DeleteHeaderSpaceLine()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_SD925H.GetValue("U_BPLId", 0)))
                {
                    errMessage = "사업장은 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_SD925H.GetValue("U_YYYYMM", 0)))
                {
                    errMessage = "전기년월은 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                // 마감일자 Check
                else if (dataHelpClass.Check_Finish_Status(oDS_PS_SD925H.GetValue("U_BPLId", 0).ToString().Trim(), oDS_PS_SD925H.GetValue("U_YYYYMM", 0).ToString().Trim()) == false)
                {
                    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 전기년월을 확인하고, 회계부서로 문의하세요.";
                    throw new Exception();
                }
                else if (Convert.ToDouble(oDS_PS_SD925H.GetValue("U_TgtWgt", 0)) <= 0)
                {
                    errMessage = "조회 중량은 필수입력사항입니다. 확인하세요.";
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// 필수입력사항 체크(라인)
        /// </summary>
        /// <returns></returns>
        private bool PS_SD925_DeleteMatrixSpaceLine()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (oMat01.VisualRowCount == 0) //라인
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PS_SD925_LoadData()
        {
            int sumWgt = 0;
            int TotalAmt = 0;
            double targetWgt;
            string sQry;
            string YYYYMM;
            string BPLId;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                YYYYMM = oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim();
                targetWgt = Convert.ToDouble(oForm.Items.Item("TgtWgt").Specific.Value.ToString().Trim());

                sQry = "SELECT U_Minor  FROM [@PS_SY001L] WHERE code = 'S011' and U_UseYN = 'Y'";
                oRecordSet01.DoQuery(sQry);
                oDS_PS_SD925H.SetValue("U_ChgPrice", 0, oRecordSet01.Fields.Item(0).Value.ToString().Trim());

                sQry = "EXEC [PS_SD925_01] '" + BPLId + "','" + YYYYMM + "'," +  targetWgt;
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_SD925L.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다.확인하세요.";
                    throw new Exception();
                }

                for (int i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_SD925L.Size)
                    {
                        oDS_PS_SD925L.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_SD925L.Offset = i;
                    oDS_PS_SD925L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_SD925L.SetValue("U_Check", i, oRecordSet01.Fields.Item("Chk").Value.ToString().Trim());
                    oDS_PS_SD925L.SetValue("U_ItemType", i, oRecordSet01.Fields.Item("ItemType").Value.ToString().Trim());
                    oDS_PS_SD925L.SetValue("U_ODLNDoc", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());
                    oDS_PS_SD925L.SetValue("U_DLN1Line", i, oRecordSet01.Fields.Item("LineNum").Value.ToString().Trim());
                    oDS_PS_SD925L.SetValue("U_ItemCode", i, oRecordSet01.Fields.Item("ItemCode").Value.ToString().Trim());
                    oDS_PS_SD925L.SetValue("U_ItemName", i, oRecordSet01.Fields.Item("Dscription").Value.ToString().Trim());
                    oDS_PS_SD925L.SetValue("U_Quantity", i, oRecordSet01.Fields.Item("Quantity").Value.ToString().Trim());
                    oDS_PS_SD925L.SetValue("U_Price", i, oRecordSet01.Fields.Item("Price").Value.ToString().Trim());
                    oDS_PS_SD925L.SetValue("U_LineTotal", i, oRecordSet01.Fields.Item("LineTotal").Value.ToString().Trim());

                    if(oRecordSet01.Fields.Item("Chk").Value.ToString().Trim() == "Y")
                    {
                        sumWgt += Convert.ToDouble(oRecordSet01.Fields.Item("Quantity").Value.ToString().Trim());
                    }
                    TotalAmt += Convert.ToDouble(oRecordSet01.Fields.Item("Quantity").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oDS_PS_SD925H.SetValue("U_SumWgt", 0, sumWgt.ToString());
                oDS_PS_SD925H.SetValue("U_TotalWgt", 0, TotalAmt.ToString());
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PS_SD925_Change_Price()
        {
            int i;
            double TotalAmt = 0;
            string sQry;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("처리중", 0, false);
                for (i = 1; i <= oMat01.RowCount; i++)
                {
                    if (oMat01.Columns.Item("Check").Cells.Item(i).Specific.Checked == true)
                    {
                        sQry = "Exec [PS_SD925_02] '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "','";
                        sQry += oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim() + "','";
                        sQry += oMat01.Columns.Item("ODLNDoc").Cells.Item(i).Specific.Value + "','";
                        sQry += oMat01.Columns.Item("DLN1Line").Cells.Item(i).Specific.Value + "',";
                        sQry += oForm.Items.Item("ChgPrice").Specific.Value.ToString().Trim() ;
                        oRecordSet01.DoQuery(sQry);

                        TotalAmt += Convert.ToDouble(oForm.Items.Item("ChgPrice").Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.Value);
                    }
                    else
                    {
                        TotalAmt += Convert.ToDouble(oMat01.Columns.Item("LineTotal").Cells.Item(i).Specific.Value);
                    }
                }
                oDS_PS_SD925H.SetValue("U_TotalAmt", 0, TotalAmt.ToString());
                sQry = "UPDATE [@PS_SD925H] SET U_Status ='Y', U_TotalAmt = '"+ TotalAmt + "' WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'";
                oRecordSet01.DoQuery(sQry);
                PSH_Globals.SBO_Application.MessageBox("변경완료");
                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PS_SD925_ReChange_Price()
        {
            int i;
            double TotalAmt = 0;
            string sQry;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("처리중", 0, false);

                for (i = 1; i <= oMat01.RowCount; i++)
                {
                    if (oMat01.Columns.Item("Check").Cells.Item(i).Specific.Checked == true)
                    {
                        sQry = "Exec [PS_SD925_02] '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "','";
                        sQry += oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim() + "','";
                        sQry += oMat01.Columns.Item("ODLNDoc").Cells.Item(i).Specific.Value + "','";
                        sQry += oMat01.Columns.Item("DLN1Line").Cells.Item(i).Specific.Value + "',";
                        sQry += oMat01.Columns.Item("Price").Cells.Item(i).Specific.Value;
                        oRecordSet01.DoQuery(sQry);
                    }
                    TotalAmt += Convert.ToDouble(oMat01.Columns.Item("LineTotal").Cells.Item(i).Specific.Value);
                }
                sQry = "UPDATE [@PS_SD925H] SET U_Status ='C', U_TotalAmt = '" + TotalAmt + "' WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'";
                oRecordSet01.DoQuery(sQry);
                PSH_Globals.SBO_Application.MessageBox("취소완료");
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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
            int i;
            double SumWgt = 0;
            string DocEntry;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_SD925_DeleteHeaderSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_SD925_DeleteMatrixSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
                            for (i = 1; i <= oMat01.VisualRowCount; i++) // 추가 및 업데이트시 선택 중량 재계산
                            {
                                if (oMat01.Columns.Item("Check").Cells.Item(i).Specific.Checked == true)
                                {
                                    SumWgt += Convert.ToDouble(oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.Value);
                                }
                            }
                            oDS_PS_SD925H.SetValue("U_SumWgt", 0, SumWgt.ToString());
                        }
                    }
                    else if (pVal.ItemUID == "ChgPriBt")
                    {
                        if (Convert.ToDouble(oDS_PS_SD925H.GetValue("U_TgtWgt", 0)) != Convert.ToDouble(oDS_PS_SD925H.GetValue("U_SumWgt", 0)))
                        {
                            errMessage = "조회 중량과 선택중량은 동일해야합니다..";
                            throw new Exception();
                        }
                        DocEntry = oForm.Items.Item("DocEntry").Specific.Value;
                        PS_SD925_Change_Price();
                        oForm.Mode = BoFormMode.fm_FIND_MODE;
                        PS_SD925_FormItemEnabled();
                        oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                        oForm.Items.Item("1").Click(BoCellClickType.ct_Regular);
                    }
                    else if (pVal.ItemUID == "ReChgPri")
                    {
                        DocEntry = oForm.Items.Item("DocEntry").Specific.Value;
                        PS_SD925_ReChange_Price();
                        oForm.Mode = BoFormMode.fm_FIND_MODE;
                        PS_SD925_FormItemEnabled();
                        oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                        oForm.Items.Item("1").Click(BoCellClickType.ct_Regular);
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        if (PS_SD925_DeleteHeaderSpaceLine() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        PS_SD925_LoadData();
                    }
                }
                PS_SD925_FormItemEnabled();
            }
            catch (Exception ex)
            {
                if(errMessage != string.Empty)
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
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
                }
                oForm.Freeze(false);
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            double SumWgt;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "Check")
                        {
                            if (oMat01.Columns.Item("Check").Cells.Item(pVal.Row).Specific.Checked == true)
                            {
                                SumWgt = Convert.ToDouble(oForm.Items.Item("SumWgt").Specific.Value);
                                SumWgt += Convert.ToDouble(oMat01.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value);

                                oDS_PS_SD925H.SetValue("U_SumWgt", 0, SumWgt.ToString());
                            }
                            else
                            {
                                SumWgt = Convert.ToDouble(oForm.Items.Item("SumWgt").Specific.Value);
                                SumWgt -= Convert.ToDouble(oMat01.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value);

                                oDS_PS_SD925H.SetValue("U_SumWgt", 0, SumWgt.ToString());
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
                    oMat01.AutoResizeColumns();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD925H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD925L);
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
                            break;
                        case "1281": //찾기
                            PS_SD925_FormItemEnabled();
                            break;
                        case "1282": //추가
                            PS_SD925_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_SD925_FormItemEnabled();
                            break;
                        case "1287": //복제
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