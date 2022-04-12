using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 통합구매 수동품의작성
    /// </summary>
    internal class PS_MM015 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.DBDataSource oDS_PS_MM015H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_MM015L; //등록라인
        public SAPbouiCOM.Grid oGrid01;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM015.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM015_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM015");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_MM015_CreateItems();
                PS_MM015_ComboBox_Setting();

                oForm.EnableMenu("1283", false); //삭제
                oForm.EnableMenu("1287", false); //복제
                oForm.EnableMenu("1286", false); //닫기
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", false); //행삭제
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
        private void PS_MM015_CreateItems()
        {
            try
            {
                oDS_PS_MM015H = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oDS_PS_MM015L = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");
                
                // 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat02 = oForm.Items.Item("Mat02").Specific;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_MM015_ComboBox_Setting()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }
                oForm.Items.Item("BPLId").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
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
        /// FormResize
        /// </summary>
        private void PS_MM015_FormResize()
        {
            try
            {
                oForm.Items.Item("Mat01").Top = (oForm.Height / 2) - 15;
                oForm.Items.Item("Mat01").Height = (oForm.Height / 2) - 38;
                oForm.Items.Item("Mat01").Left = 6;
                oForm.Items.Item("Mat01").Width = oForm.Width - 21;

                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_MM015_Search_Data
        /// </summary>
        private void PS_MM015_Search_Data()
        {
            string errMessage = string.Empty;
            int i;
            string sQry;
            string BPLId;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                oMat01.Clear();
                BPLId = oForm.Items.Item("BPLId").Specific.Selected.VALUE.ToString().Trim();

                sQry = "EXEC PS_MM015_01 '" + BPLId + "'";
                oRecordSet.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_MM015H.Clear();
                if (oRecordSet.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다. 확인하세요.";
                    throw new Exception();
                }
                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_MM015H.Size)
                    {
                        oDS_PS_MM015H.InsertRecord(i);
                    }
                    oMat01.AddRow();
                    oDS_PS_MM015H.Offset = i;
                    oDS_PS_MM015H.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("BEDAT").Value.ToString().Trim());
                    oDS_PS_MM015H.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("EBELN").Value.ToString().Trim());
                    oDS_PS_MM015H.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("Lifnr").Value.ToString().Trim());
                    oDS_PS_MM015H.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("CardCode").Value.ToString().Trim());
                    oDS_PS_MM015H.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("CardName").Value.ToString().Trim());
                    oDS_PS_MM015H.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("Cnt").Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();

                oMat02.Clear();
                oForm.Items.Item("S_MENGE").Specific.Value = "";
                oForm.Items.Item("S_NETWR").Specific.Value = "";
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_MM015_Search_Matrix_Data
        /// </summary>
        private void PS_MM015_Search_Matrix_Data(int ClickRow)
        {
            string errMessage = string.Empty;
            int j;
            int cnt;
            string EBELN;
            string sQry;
            double S_MENGE = 0; //품의중량
            double S_NETWR = 0; //품의금액
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                EBELN = oMat01.Columns.Item("EBELN").Cells.Item(ClickRow).Specific.VALUE.ToString().Trim();
               
                sQry = "EXEC PS_MM015_02 '" + EBELN + "'";
                oRecordSet.DoQuery(sQry);
                
                cnt = oDS_PS_MM015L.Size;

                if (cnt > 0)
                {
                    for (j = 0; j <= cnt - 1; j++)
                    {
                        oDS_PS_MM015L.RemoveRecord(oDS_PS_MM015L.Size - 1);
                    }
                    if (cnt == 1)
                    {
                        oDS_PS_MM015L.Clear();
                    }
                }
                oMat02.LoadFromDataSource();

                //Matrix에 Data 뿌려준다
                j = 1;
                while (!oRecordSet.EoF)
                {
                    if (oDS_PS_MM015L.Size < j)
                    {
                        oDS_PS_MM015L.InsertRecord(j - 1); //라인추가
                    }
                    oDS_PS_MM015L.SetValue("U_ColReg01", j - 1, oRecordSet.Fields.Item("BPLId").Value);
                    oDS_PS_MM015L.SetValue("U_ColReg02", j - 1, oRecordSet.Fields.Item("DocNum").Value);
                    oDS_PS_MM015L.SetValue("U_ColReg03", j - 1, oRecordSet.Fields.Item("LineNum").Value);
                    oDS_PS_MM015L.SetValue("U_ColReg04", j - 1, oRecordSet.Fields.Item("CardCode").Value);
                    oDS_PS_MM015L.SetValue("U_ColReg05", j - 1, oRecordSet.Fields.Item("Purchase").Value);
                    oDS_PS_MM015L.SetValue("U_ColReg06", j - 1, oRecordSet.Fields.Item("PQType").Value);
                    oDS_PS_MM015L.SetValue("U_ColReg07", j - 1, oRecordSet.Fields.Item("itemCode").Value);
                    oDS_PS_MM015L.SetValue("U_ColQTy01", j - 1, oRecordSet.Fields.Item("PQTy").Value);
                    oDS_PS_MM015L.SetValue("U_ColQTy02", j - 1, oRecordSet.Fields.Item("Weight").Value);
                    oDS_PS_MM015L.SetValue("U_ColQTy03", j - 1, oRecordSet.Fields.Item("MENGE").Value); //품의수량
                    oDS_PS_MM015L.SetValue("U_ColQTy04", j - 1, oRecordSet.Fields.Item("NETWR").Value);  //품의금액
                    oDS_PS_MM015L.SetValue("U_ColReg08", j - 1, oRecordSet.Fields.Item("CntcCode").Value);
                    oDS_PS_MM015L.SetValue("U_ColReg09", j - 1, oRecordSet.Fields.Item("BEDAT").Value);
                    oDS_PS_MM015L.SetValue("U_ColReg10", j - 1, oRecordSet.Fields.Item("EINDT").Value);

                    S_MENGE += oRecordSet.Fields.Item("MENGE").Value;
                    S_NETWR += oRecordSet.Fields.Item("NETWR").Value;
                    j += 1;
                    oRecordSet.MoveNext();
                }
                oMat02.LoadFromDataSource();
                
                oForm.Items.Item("S_MENGE").Specific.Value = S_MENGE;
                oForm.Items.Item("S_NETWR").Specific.Value = S_NETWR;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }


        /// <summary>
        /// PS_MM015_Save_Data
        /// </summary>
        private bool PS_MM015_Save_Data()
        {
            bool ReturnValue = false;
            int RetVal;
            int i;
            string sQry;
            string CntcCode;
            string BPLId;
            string LineNum;
            string PODocEntry;
            string DocNum;
            string CardCode;
            string DocDate;
            string DueDate;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Documents DI_oPurchaseOrders = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

            try
            {
                oMat02.FlushToDataSource();
                if (oMat02.VisualRowCount > 0)
                {
                    i = 0;
                    oDS_PS_MM015L.Offset = i;
                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }
                    PSH_Globals.oCompany.StartTransaction();

                    DocNum = oDS_PS_MM015L.GetValue("U_ColReg02", i).ToString().Trim();
                    LineNum = oDS_PS_MM015L.GetValue("U_ColReg03", i).ToString().Trim();
                    CardCode = oDS_PS_MM015L.GetValue("U_ColReg04", i).ToString().Trim();
                    BPLId = oDS_PS_MM015L.GetValue("U_ColReg01", i).ToString().Trim();
                    DocDate = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oDS_PS_MM015L.GetValue("U_ColReg09", i).ToString().Trim(), "YYYY-MM-DD");
                    DueDate = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oDS_PS_MM015L.GetValue("U_ColReg10", i).ToString().Trim(), "YYYY-MM-DD");
                    CntcCode = oDS_PS_MM015L.GetValue("U_ColReg08", i).ToString().Trim();
                    
                    DI_oPurchaseOrders.CardCode = CardCode;
                    DI_oPurchaseOrders.BPL_IDAssignedToInvoice = Convert.ToInt32(BPLId);
                    DI_oPurchaseOrders.DocDate = Convert.ToDateTime(DocDate);
                    DI_oPurchaseOrders.DocDueDate = Convert.ToDateTime(DueDate);

                    sQry = "Select empID From OHEM Where U_MSTCOD = '" + CntcCode + "'";
                    oRecordSet.DoQuery(sQry);

                    DI_oPurchaseOrders.DocumentsOwner = Convert.ToInt32(oRecordSet.Fields.Item("empID").Value).ToString().Trim();
                    DI_oPurchaseOrders.UserFields.Fields.Item("U_reType").Value = oDS_PS_MM015L.GetValue("U_ColReg06", i).ToString().Trim();
                    DI_oPurchaseOrders.UserFields.Fields.Item("U_okYN").Value = "N";
                    DI_oPurchaseOrders.UserFields.Fields.Item("U_OrdTyp").Value = oDS_PS_MM015L.GetValue("U_ColReg05", i).ToString().Trim();
                    DI_oPurchaseOrders.Lines.SetCurrentLine(i);
                    DI_oPurchaseOrders.Lines.ItemCode = oDS_PS_MM015L.GetValue("U_ColReg07", i).ToString().Trim();
                    DI_oPurchaseOrders.Lines.Quantity = Convert.ToDouble(oDS_PS_MM015L.GetValue("U_ColQTy03", i).ToString().Trim());
                    DI_oPurchaseOrders.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM015L.GetValue("U_ColQTy04", i).ToString().Trim());
                    DI_oPurchaseOrders.Lines.WarehouseCode = "10" + BPLId;
                    DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_Qty").Value = oDS_PS_MM015L.GetValue("U_ColQTy03", i).ToString().Trim();
                    DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_MM010Doc").Value = DocNum;
                    DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_MM010Lin").Value = LineNum;
                    DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_Auto").Value = "N";

                    if (oMat02.VisualRowCount > 1)
                    {
                        for (i = 1; i <= oMat02.VisualRowCount - 1; i++)
                        {
                            DocNum = oDS_PS_MM015L.GetValue("U_ColReg02", i).ToString().Trim();
                            LineNum = oDS_PS_MM015L.GetValue("U_ColReg03", i).ToString().Trim();

                            if (i > 0)
                            {
                                DI_oPurchaseOrders.Lines.Add();
                            }
                            DI_oPurchaseOrders.Lines.ItemCode = oDS_PS_MM015L.GetValue("U_ColReg07", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.Quantity = Convert.ToDouble(oDS_PS_MM015L.GetValue("U_ColQTy03", i).ToString().Trim());
                            DI_oPurchaseOrders.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM015L.GetValue("U_ColQTy04", i).ToString().Trim());
                            DI_oPurchaseOrders.Lines.WarehouseCode = "10" + BPLId;
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_Qty").Value = oDS_PS_MM015L.GetValue("U_ColQTy03", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_MM010Doc").Value = DocNum;
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_MM010Lin").Value = LineNum;
                        }
                    }

                    RetVal = DI_oPurchaseOrders.Add();
                    if (RetVal != 0)
                    {
                        if (PSH_Globals.oCompany.InTransaction == true)
                        {
                            PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        }
                    }
                    else
                    {
                        PSH_Globals.oCompany.GetNewObjectCode(out PODocEntry);
                        sQry = "EXEC [PS_INTERFACE_01] '" + PODocEntry + "'";
                        oRecordSet.DoQuery(sQry);
                        
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                        oMat02.Clear();
                        oForm.Items.Item("S_MENGE").Specific.Value = "";
                        oForm.Items.Item("S_NETWR").Specific.Value = "";

                        sQry = "Update [@PS_MM010L] Set  U_GuBun = '3' ";
                        sQry += "From  [@PS_MM010L] a Inner Join [@PS_Mm010H] b On a.DocEntry = b.DocEntry ";
                        sQry += "Where a.U_GuBun = '2' And  b.U_PQType = '20' And b.CanCeled = 'N' And Isnull(a.U_POYesNo, 'N') = 'Y'";
                        oRecordSet.DoQuery(sQry);
                        
                        PS_MM015_Search_Data();
                    }
                }
                ReturnValue = true;
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
                oForm.Freeze(false);
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

                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                    if (pVal.ItemUID == "Search")
                    {
                        PS_MM015_Search_Data(); //선택저장버튼
                    }
                    else if (pVal.ItemUID == "Save")
                    {
                        if (Convert.ToInt32(PSH_Globals.SBO_Application.MessageBox("이 데이터를 추가한 후에는 변경할 수 없습니다. 계속하겠습니까?", 1, "&확인", "&취소")) == 1)
                        {
                            if (PS_MM015_Save_Data() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
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
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {

                        PS_MM015_Search_Matrix_Data(pVal.Row);
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// RESIZE 이벤트
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
                    PS_MM015_FormResize();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM015H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM015L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
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
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
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
            finally
            {
            }
        }
    }
}
