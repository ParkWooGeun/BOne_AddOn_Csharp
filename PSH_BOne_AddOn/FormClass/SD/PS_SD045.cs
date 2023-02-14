using System;
using System.Linq;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using System.Collections.Generic;
using PSH_BOne_AddOn.Code;
using System.Timers;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using SAP.Middleware.Connector;


namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 반품처리(이월) 
    /// </summary>
    internal class PS_SD045 : PSH_BaseClass
    {
        public string SDocEntry = null;
        public string ORDNDocEnt = null;
        public string sVersion = null;
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Grid oGrid01;
        private SAPbouiCOM.DBDataSource oDS_PS_SD045H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_SD045L; //등록라인
        private SAPbouiCOM.DataTable oDS_PS_SD045C;
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD045.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_SD045_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_SD045");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_SD045_CreateItems();
                PS_SD045_FormItemEnabled();
                PS_SD045_SetDocEntry();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>        
        private void PS_SD045_CreateItems()
        {
            try
            {
                oDS_PS_SD045H = oForm.DataSources.DBDataSources.Item("@PS_SD045H");
                oDS_PS_SD045L = oForm.DataSources.DBDataSources.Item("@PS_SD045L");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMat01.AutoResizeColumns();

                oGrid01 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PS_SD045C");
                oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_SD045C");
                oDS_PS_SD045C = oForm.DataSources.DataTables.Item("PS_SD045C");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_SD045_SetComboBox()
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
        private void PS_SD045_SetInitial()
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
        /// DocEntry 초기화
        /// </summary>
        private void PS_SD045_SetDocEntry()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD045'", "");
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
        private void PS_SD045_FormItemEnabled()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("FrDt").Enabled = true;
                    oForm.Items.Item("ToDt").Enabled = true;
                    oForm.Items.Item("SDocNum").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                    oForm.Items.Item("btn_srh").Enabled = true;
                    oMat01.Columns.Item("Weight").Editable = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("FrDt").Enabled = true;
                    oForm.Items.Item("ToDt").Enabled = true;
                    oForm.Items.Item("SDocNum").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                    oForm.Items.Item("btn_srh").Enabled = false;
                    oMat01.Columns.Item("Weight").Editable = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("FrDt").Enabled = false;
                    oForm.Items.Item("ToDt").Enabled = false;
                    oForm.Items.Item("SDocNum").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가    
                    oForm.Items.Item("btn_srh").Enabled = false;
                    oMat01.Columns.Item("Weight").Editable = false;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스 행추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_SD045_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_SD045L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_SD045L.Offset = oRow;
                oDS_PS_SD045L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private bool PS_SD045_DeleteHeaderSpaceLine()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("SDocNum").Specific.Value.ToString().Trim())) //|| string.IsNullOrEmpty(oDS_PS_SD045H.GetValue("U_ToDt", 0)))
                {
                    errMessage = "납품문서번호는 필수입력사항입니다. 확인하세요.";
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
        /// 납품처리문서 데이터 조회 
        /// </summary>
        private bool LoadData01(string DocEntry)
        {
            bool returnValue = false;
            string sQry;
            int i = 0;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                sQry = "EXEC [PS_SD045_01] '"+ DocEntry + "'";
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_SD045L.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();


                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다.확인하세요.";
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_SD045L.InsertRecord(i);
                    }

                    oDS_PS_SD045L.Offset = i;
                    oDS_PS_SD045L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_SD045L.SetValue("U_BatchNum", i, oRecordSet01.Fields.Item("U_BatchNum").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_chWeight", i, oRecordSet01.Fields.Item("U_chWeight").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_CardCode", i, oRecordSet01.Fields.Item("U_CardCode").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_CardName", i, oRecordSet01.Fields.Item("U_CardName").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_ItemCode", i, oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_ItemName", i, oRecordSet01.Fields.Item("U_ItemName").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_Qty", i, oRecordSet01.Fields.Item("U_Qty").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_Price", i, oRecordSet01.Fields.Item("U_Price").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_LineTotal", i, oRecordSet01.Fields.Item("U_LinTotal").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_WhsCod", i, oRecordSet01.Fields.Item("U_WhsCode").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_BinCod", i, oRecordSet01.Fields.Item("U_WhsName").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_SD040Doc", i, oRecordSet01.Fields.Item("U_SD040Doc").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_SD040Lin", i, oRecordSet01.Fields.Item("U_SD040Lin").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_ORDRDoc", i, oRecordSet01.Fields.Item("U_ORDRDoc").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_RDR1Line", i, oRecordSet01.Fields.Item("U_RDR1Line").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_ODLNDoc", i, oRecordSet01.Fields.Item("U_ODLNDoc").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_DLNLine", i, oRecordSet01.Fields.Item("U_DLNLine").Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

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
            return returnValue;
        }

        /// <summary>
        /// 반품
        /// </summary>
        /// <returns></returns>
        private bool PS_SD045_DI_API()
        {
            bool returnValue = false;
            string errMsg = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int i;
            int RetVal;
            double Remain = 0;
            int LineNumCount;
            double OWeight = 0;
            string Query01;
            string lclDocCur;
            string lclQuery;
            double lclDocRate;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Documents oDIObject = null;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset RecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errMsg = "현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.";
                    throw new Exception();
                }

                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();

                oMat01.FlushToDataSource();
                SDocEntry = oForm.Items.Item("SDocNum").Specific.Value.ToString().Trim(); //납품문서번호 미리저장용)

                LineNumCount = 0;
                Query01 = "  SELECT U_BPLID,U_CardCode,U_DCardCod,U_DCardNam,U_TrType FROM [@PS_SD040H] ";
                Query01 +=    "WHERE DocEntry = '" + oForm.Items.Item("SDocNum").Specific.Value.ToString().Trim() + "'";
                RecordSet01.DoQuery(Query01);

                oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns);
                oDIObject.BPL_IDAssignedToInvoice = Convert.ToInt32(RecordSet01.Fields.Item("U_BPLId").Value);
                oDIObject.CardCode = RecordSet01.Fields.Item("U_CardCode").Value;
                oDIObject.UserFields.Fields.Item("U_DCardCod").Value = RecordSet01.Fields.Item("U_DCardCod").Value;
                oDIObject.UserFields.Fields.Item("U_DCardNam").Value = RecordSet01.Fields.Item("U_DCardNam").Value;
                oDIObject.UserFields.Fields.Item("U_TradeType").Value = RecordSet01.Fields.Item("U_TrType").Value;
                oDIObject.DocDate = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd"));
                oDIObject.DocDueDate = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd"));
              
                //헤더 환율처리_S (2011.11.04 송명규 추가)
                lclQuery = "  SELECT  Currency,";
                lclQuery += "         Rate";
                lclQuery += " FROM    DLN1 ";
                lclQuery += " WHERE   DocEntry = " + oDS_PS_SD045L.GetValue("U_ODLNDoc",1).ToString().Trim() + " ";
                lclQuery += "         AND LineNum = " + oDS_PS_SD045L.GetValue("U_DLNLine", 1);
                RecordSet02.DoQuery(lclQuery);

                lclDocCur = RecordSet02.Fields.Item("Currency").Value.ToString().Trim();
                lclDocRate = Convert.ToDouble(RecordSet02.Fields.Item("Rate").Value.ToString().Trim());

                if (lclDocCur != "KRW")
                {
                    oDIObject.DocCurrency = lclDocCur;
                    oDIObject.DocRate = lclDocRate;
                }
                //헤더 환율처리_E

                for (i = 0; i <= oMat01.VisualRowCount-1; i++)
                {
                    if (!string.IsNullOrEmpty((oDS_PS_SD045L.GetValue("U_Weight", i).ToString().Trim())))
                    {
                        OWeight = Convert.ToDouble(oDS_PS_SD045L.GetValue("U_Weight", i).ToString().Trim());
                        Remain = Convert.ToDouble(oDS_PS_SD045L.GetValue("U_ChWeight", i).ToString().Trim());

                        oDIObject.Lines.Add();
                        oDIObject.Lines.SetCurrentLine(LineNumCount);
                        oDIObject.Lines.ItemCode = oDS_PS_SD045L.GetValue("U_ItemCode", i).ToString().Trim();
                        oDIObject.Lines.WarehouseCode = oDS_PS_SD045L.GetValue("U_WhsCod", i).ToString().Trim();
                        oDIObject.Lines.UserFields.Fields.Item("U_BaseType").Value = "PS_SD040";
                        oDIObject.Lines.UserFields.Fields.Item("U_Qty").Value = oDS_PS_SD045L.GetValue("U_Qty", i).ToString().Trim();
                        oDIObject.Lines.BaseType = 15;
                        oDIObject.Lines.BaseEntry = Convert.ToInt32(oDS_PS_SD045L.GetValue("U_ODLNDoc", i).ToString().Trim());
                        oDIObject.Lines.BaseLine = Convert.ToInt32(oDS_PS_SD045L.GetValue("U_DLNLine", i).ToString().Trim());
                        oDIObject.Lines.Quantity = OWeight; //Convert.ToDouble(oMat01.Columns.Item("OWeight1").Cells.Item(i).Specific.Value.ToString().Trim());

                        if (Remain < OWeight)
                        {
                            errMsg = "반품수량을 초과하였습니다. 다시 확인해주세요.";
                            throw new Exception();
                        }
                        if (lclDocCur != "KRW")
                        
                        {
                            oDIObject.Lines.Currency = lclDocCur;
                            oDIObject.Lines.Rate = lclDocRate;
                        }

                        if (dataHelpClass.GetItem_ManBtchNum(Convert.ToString(oDS_PS_SD045L.GetValue("U_ItemCode", i).ToString().Trim())) == "Y")
                        {
                            oDIObject.Lines.BatchNumbers.BatchNumber = oDS_PS_SD045L.GetValue("U_BatchNum", i).ToString().Trim();
                            oDIObject.Lines.BatchNumbers.Quantity = OWeight;
                            oDIObject.Lines.BatchNumbers.Add();
                        }
                        LineNumCount += 1;
                    }
                }
                RetVal = oDIObject.Add();

                if (RetVal == 0)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out string afterDIDocNum);
                    ORDNDocEnt = afterDIDocNum;

                    Query01 = "EXECUTE [PS_Z_RETU_GI_ZSD920] '" + afterDIDocNum + "'";
                    RecordSet01.DoQuery(Query01);
                   
                }
                else
                {
                    PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                    errMsg = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                    throw new Exception();
                }
                PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errMsg != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMsg);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                if (oDIObject != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
                }

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet02);
            }

            return returnValue;
        }



        /// <summary>
        /// SD920L 반품, 납품 문서번호 업데이트
        /// </summary>
        private bool PS_SD045_UpdateData(string DocEntry)
        {
            bool returnValue = false;
            int i;
            string Query01;
            string Query02;
            string SDocNum;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset RecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                int a = 0;
                Query01 = "Select * from [@PS_SD045L] WHERE DocEntry = '"+ DocEntry +"'";
                RecordSet01.DoQuery(Query01);
                SDocNum = RecordSet01.Fields.Item("U_SD040Doc").Value.ToString().Trim();
                
                for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                {
                    if ((RecordSet01.Fields.Item("U_Weight").Value) > 0)
                    {
                        Query02 = " UPDATE [@PS_SD040L]";
                        Query02 += " SET U_ORDNNum = '" + ORDNDocEnt + "', U_RDN1Num = '" + a + "'";
                        Query02 += "WHERE U_LotNo ='" + RecordSet01.Fields.Item("U_BatchNum").Value.ToString().Trim() + "'  AND DocEntry ='" + RecordSet01.Fields.Item("U_SD040Doc").Value.ToString().Trim() + "'";
                        RecordSet02.DoQuery(Query02);

                        Query02 = " UPDATE [@PS_SD045L]";
                        Query02 += " SET U_ORDNDoc = '" + ORDNDocEnt + "', U_RDN1Line = '" + a + "'";
                        Query02 += "WHERE U_BatchNum ='" + RecordSet01.Fields.Item("U_BatchNum").Value.ToString().Trim() + "'  AND DocEntry ='" + DocEntry + "'";
                        RecordSet02.DoQuery(Query02);
                        a += 1;
                    }
                    RecordSet01.MoveNext();
                }

                Query01 = " Select Count(*) from [@PS_SD040L] WHERE U_ORDNNum IS NULL AND DocEntry = '" + oForm.Items.Item("SDocNum").Specific.Value.ToString().Trim() + "'";
                RecordSet01.DoQuery(Query01);

                if (Convert.ToDouble(RecordSet01.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    dataHelpClass.DoQuery("UPDATE [@PS_SD040H] SET U_ProgStat = '4', Canceled ='Y', Status = 'C' WHERE DocEntry = '" + SDocNum + "'");
                    Query01 = "SELECT Convert(CHAR(8),U_DocDate,112) AS U_DocDate, U_CardCode FROM [@PS_SD040H] WHERE DocEntry ='" + SDocNum + "'";
                    RecordSet01.DoQuery(Query01);
                    string DocDate2 = RecordSet01.Fields.Item("U_DocDate").Value;
                    string CardCode2 = RecordSet01.Fields.Item("U_CardCode").Value;

                    //여신한도초과요청:납품처리여부 필드 업데이트(KEY-해당일자, 거래처코드) 
                    Query01 = "  UPDATE	[@PS_SD080L]";
                    Query01 += " SET    U_SD040YN = 'N'"; //납품처리여부 "N"으로 환원
                    Query01 += "        FROM[@PS_SD080H] AS T0";
                    Query01 += "        INNER JOIN";
                    Query01 += "        [@PS_SD080L] AS T1";
                    Query01 += "            ON T0.DocEntry = T1.DocEntry";
                    Query01 += " WHERE  T0.U_DocDate = '" + DocDate2 + "'"; //해당일자
                    Query01 += "        AND T1.U_CardCode = '" + CardCode2 + "'"; //해당거래처코드
                    RecordSet01.DoQuery(Query01);
                }

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                PS_SD045_FormItemEnabled();
                oForm.Items.Item("DocEntry").Specific.Value = sVersion;
                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                returnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                
                if (pVal.BeforeAction == true)
                {
                    sVersion = oForm.Items.Item("DocEntry").Specific.Value;
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            oDS_PS_SD045H.SetValue("U_UPName", 0, dataHelpClass.User_MSTCOD()); //담당자
                            
                            if (PS_SD045_DI_API() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                    }
                    if (pVal.ItemUID == "btn_srh")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_SD045_DeleteHeaderSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (LoadData01(oForm.Items.Item("SDocNum").Specific.Value) == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                    } }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
                        {
                            if (PS_SD045_UpdateData(sVersion) == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PS_SD045_FormItemEnabled();
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
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string ColReg01;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01" && pVal.ColUID == "SD040Doc")
                    {
                        ColReg01 = oMat01.Columns.Item("SD040Doc").Cells.Item(pVal.Row).Specific.Value;
                        PS_SD040 pS_SD040 = new PS_SD040();
                        pS_SD040.LoadForm(ColReg01);
                        pS_SD040 = null;
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "CardCode":
                                sQry = "SELECT CardFName FROM OCRD WHERE CardCode =  '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item("CardFName").Value.ToString().Trim();
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
        /// Raise_EVENT_DOUBLE_CLICK
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row == 0)
                        {
                            oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat01.FlushToDataSource();
                        }
                        else
                        {
                        }
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
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "CardCode" || pVal.ItemUID == "CardName")
                    {
                        dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_SD045H", "U_CardCode,U_CardName", "", 0, "", "", "");
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
                        if (pVal.ItemUID == "CardCode")
                        {
                            if (oForm.Items.Item("CardCode").Specific.Value == "")
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (pVal.ItemUID == "SDocNum")
                        {
                            if (oForm.Items.Item("SDocNum").Specific.Value == "")
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                                return;
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
                    PS_SD045_AddMatrixRow(oMat01.RowCount, false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD045H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD045L);
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
                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
                                }

                                oMat01.FlushToDataSource();
                                oDS_PS_SD045L.RemoveRecord(oDS_PS_SD045L.Size - 1);
                                oMat01.Clear();
                                oMat01.LoadFromDataSource();

                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("LineNum").Cells.Item(oMat01.RowCount).Specific.Value))
                                {
                                    PS_SD045_AddMatrixRow(oMat01.RowCount, false);
                                }
                            }
                            break;
                        case "1281": //찾기
                            PS_SD045_FormItemEnabled();
                            break;
                        case "1282": //추가
                            oDS_PS_SD045H.SetValue("U_UPName", 0, dataHelpClass.User_MSTCOD()); //담당자
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
                                oDS_PS_SD045L.RemoveRecord(oDS_PS_SD045L.Size - 1);
                                oMat01.Clear();
                                oMat01.LoadFromDataSource();

                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("LineNum").Cells.Item(oMat01.RowCount).Specific.Value))
                                {
                                    PS_SD045_AddMatrixRow(oMat01.RowCount, false);
                                }
                            }
                            break;
                        case "1281": //찾기
                            PS_SD045_FormItemEnabled();
                            break;
                        case "1282": //추가
                            PS_SD045_FormItemEnabled();
                            PS_SD045_SetDocEntry();
                            oDS_PS_SD045H.SetValue("U_UPName", 0, dataHelpClass.User_MSTCOD()); //담당자
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_SD045_FormItemEnabled();
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
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                        //Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        //Raise_EVENT_FORM_DATA_ADD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        //Raise_EVENT_FORM_DATA_UPDATE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        //Raise_EVENT_FORM_DATA_DELETE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
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