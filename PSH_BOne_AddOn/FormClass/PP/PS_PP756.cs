using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using SAP.Middleware.Connector;
using PSH_BOne_AddOn.Code;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// Lot수율본사전송
    /// </summary>
    internal class PS_PP756 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_PP756H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_PP756L; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값


        /// <summary>
        /// LoadForm
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP756.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP756_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP756");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_PP756_CreateItems();
                PS_PP756_EnableMenus();
                PS_PP756_SetDocument(oFormDocEntry);
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
        private void PS_PP756_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oDS_PS_PP756H = oForm.DataSources.DBDataSources.Item("@PS_PP756H");
                oDS_PS_PP756L = oForm.DataSources.DBDataSources.Item("@PS_PP756L");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                //사업장
                sQry = "SELECT BPLId, BPLName From[OBPL] order by 1";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("BPLId").Specific, "N");
                oForm.Items.Item("BPLId").DisplayDesc = true;

                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); //아이디별 사업장 세팅

                oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }


        /// <summary>
		/// EnableMenus
		/// </summary>
        private void PS_PP756_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, false, true, false, true, true, true, true, false, false, false, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면 초기화
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        private void PS_PP756_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_PP756_EnableFormItem();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PS_PP756_EnableFormItem();
                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
        private void PS_PP756_EnableFormItem()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    PS_PP756_SetDocEntry();
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = false;
                }
                oMat01.AutoResizeColumns();
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
        /// PS_PP756_MTX01
        /// </summary>
        private void PS_PP756_MTX01()
        {
            int i;
            string sQry;
            string errMessage = string.Empty;
            string Param01;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("DocDate").Specific.Value.Trim();

                //sQry = "SELECT COUNT(*) FROM [@PS_PP756H] WHERE U_DocDate = '" + Param01 + "'";
                //oRecordSet.DoQuery(sQry);
                //if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) > 0)
                //{
                //    errMessage = "중복값이 존재합니다.";
                //    oMat01.Clear();
                //    throw new Exception();
                //}

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                sQry = "EXEC PS_PP756_02 '" + Param01  + "'";
                oRecordSet.DoQuery(sQry);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();
                oDS_PS_PP756L.Clear(); //추가

                if (oRecordSet.RecordCount == 0)
                {
                    errMessage = "결과값이 존재하지않습니다.";
                    oMat01.Clear();
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_PP756L.Size)
                    {
                        oDS_PS_PP756L.InsertRecord(i);
                    }
                    oMat01.AddRow();

                    oDS_PS_PP756L.Offset = i;
                    oDS_PS_PP756L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_PP756L.SetValue("U_BatcHNum", i, oRecordSet.Fields.Item("U_BatcHNum").Value);
                    oDS_PS_PP756L.SetValue("U_BoxNo", i, oRecordSet.Fields.Item("U_BoxNo").Value);
                    oDS_PS_PP756L.SetValue("U_YDocDate", i, oRecordSet.Fields.Item("YDocDate").Value);
                    oDS_PS_PP756L.SetValue("U_R3PONum", i, oRecordSet.Fields.Item("U_R3PONum").Value);
                    oDS_PS_PP756L.SetValue("U_RItemCod", i, oRecordSet.Fields.Item("U_RItemCod").Value);
                    oDS_PS_PP756L.SetValue("U_ItemCode", i, oRecordSet.Fields.Item("U_ItemCode").Value);
                    oDS_PS_PP756L.SetValue("U_ItemName", i, oRecordSet.Fields.Item("U_ItemName").Value);
                    oDS_PS_PP756L.SetValue("U_EA", i, oRecordSet.Fields.Item("EA").Value);
                    oDS_PS_PP756L.SetValue("U_BQty", i, oRecordSet.Fields.Item("U_BQty").Value);
                    oDS_PS_PP756L.SetValue("U_YQty", i, oRecordSet.Fields.Item("U_YQty").Value);
                    oDS_PS_PP756L.SetValue("U_NQty", i, oRecordSet.Fields.Item("U_NQty").Value);
                    oDS_PS_PP756L.SetValue("U_Scrap", i, oRecordSet.Fields.Item("Scrap").Value);
                    oDS_PS_PP756L.SetValue("U_NEQUIP", i, oRecordSet.Fields.Item("NEQUIP").Value);
                    oDS_PS_PP756L.SetValue("U_NEDCOD1", i, oRecordSet.Fields.Item("NEDCOD1").Value);
                    oDS_PS_PP756L.SetValue("U_NEDCOD2", i, oRecordSet.Fields.Item("NEDCOD2").Value);
                    oDS_PS_PP756L.SetValue("U_FailName", i, oRecordSet.Fields.Item("U_FailName").Value);
                    oDS_PS_PP756L.SetValue("U_NDocDate", i, oRecordSet.Fields.Item("NDocDate").Value);
                    oDS_PS_PP756L.SetValue("U_suyal", i, oRecordSet.Fields.Item("suyal").Value);
                    oDS_PS_PP756L.SetValue("U_Itmsuyal", i, oRecordSet.Fields.Item("Itmsuyal").Value);
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
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PS_PP756_MTX01:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP756_SetDocEntry()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                string DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP756'", "");
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
        /// 필수 입력 사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_PP756_CheckDataValid()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    errMessage = "전기일은 필수입니다.";
                    oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                oMat01.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP756_SetDocEntry();
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
            finally
            {
            }

            return returnValue;
        }

        /// <summary>
        /// 본사 데이터 전송
        /// </summary>
        private bool PS_PP756_InterfaceB1toR3()
        {
            bool returnValue = false;
            string sQry;
            string Client; //클라이언트
            string ServerIP; //서버IP
            string errCode = string.Empty;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            RfcDestination rfcDest = null;
            RfcRepository rfcRep = null;

            try
            {
                oMat01.FlushToDataSource();

                Client = dataHelpClass.GetR3ServerInfo()[0];
                ServerIP = dataHelpClass.GetR3ServerInfo()[1];

                //0. 연결
                if (dataHelpClass.SAPConnection(ref rfcDest, ref rfcRep, "PST", ServerIP, Client, "ifuser", "pdauser") == false)
                {
                    errCode = "1";
                    throw new Exception();
                }

                //1. SAP R3 함수 호출(매개변수 전달)
                IRfcFunction oFunction = rfcRep.CreateFunction("ZPP_HOLDINGS_INTF_YIELD");
                IRfcTable oTable = oFunction.GetTable("IT_60601");
                oTable.Insert();

                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    oTable.SetValue("ZLOTCOIL", oDS_PS_PP756L.GetValue("U_BatchNum", i).ToString().Trim()); //온산LOTNo
                    oTable.SetValue("ZBOXNO", oDS_PS_PP756L.GetValue("U_BoxNo", i).ToString().Trim()); //BoxNo
                    oTable.SetValue("ERSDA", oDS_PS_PP756L.GetValue("U_YDocDate", i).ToString().Trim()); //납품일자
                    oTable.SetValue("PONO", oDS_PS_PP756L.GetValue("U_R3PONum", i).ToString().Trim()); //모재코드
                    oTable.SetValue("MATNR", oDS_PS_PP756L.GetValue("U_RItemCod", i).ToString().Trim()); //R3PONo
                    oTable.SetValue("MATNR2", oDS_PS_PP756L.GetValue("U_ItemCode", i).ToString().Trim()); //창원품목코드
                    oTable.SetValue("MAKTX2", oDS_PS_PP756L.GetValue("U_ItemName", i).ToString().Trim()); //창원품목명
                    oTable.SetValue("GEWEI", oDS_PS_PP756L.GetValue("U_EA", i).ToString().Trim()); //단위(kg)
                    oTable.SetValue("ZEXTWE", oDS_PS_PP756L.GetValue("U_BQty", i).ToString().Trim()); //투입량
                    oTable.SetValue("ZOKYWE", oDS_PS_PP756L.GetValue("U_YQty", i).ToString().Trim()); //생산량
                    oTable.SetValue("ZBDYWE", oDS_PS_PP756L.GetValue("U_NQty", i).ToString().Trim()); //불량량
                    oTable.SetValue("ZSCRWE", oDS_PS_PP756L.GetValue("U_Scrap", i).ToString().Trim()); //Scrap
                    oTable.SetValue("ZEQUIP", oDS_PS_PP756L.GetValue("U_NEQUIP", i).ToString().Trim()); //불량공정코드
                    oTable.SetValue("ZEDCOD1", oDS_PS_PP756L.GetValue("U_NEDCOD1", i).ToString().Trim()); //불량요인
                    oTable.SetValue("ZEDCOD2", oDS_PS_PP756L.GetValue("U_NEDCOD2", i).ToString().Trim()); //불량요인2
                    oTable.SetValue("ZMDATE", oDS_PS_PP756L.GetValue("U_NDocDate", i).ToString().Trim()); //불량일자
                    oTable.Append();
                    
                    if (oMat01.VisualRowCount == i + 1) //마지막에 실행
                    {
                        errCode = "2"; //SAP Function 실행 오류가 발생했을 때 에러코드로 처리하기 위해 이 위치에서 "2"를 대입
                        oFunction.Invoke(rfcDest); //Function 실행
                    }

                    if (oFunction.GetValue("E_MESSAGE").ToString().Trim() != "" && codeHelpClass.Left(oFunction.GetValue("E_MESSAGE").ToString().Trim(), 1) != "S") //리턴 메시지가 "S(성공)"이 아니면
                    {
                        errCode = "3";
                        errMessage = oFunction.GetValue("E_MESSAGE").ToString();
                        throw new Exception();
                    }
                }
                oMat01.LoadFromDataSource();
                PSH_Globals.SBO_Application.MessageBox("R3 인터페이스 완료!");
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("풍산 SAP R3에 로그온 할 수 없습니다. 관리자에게 문의 하세요.");
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox("RFC Function 호출 오류");
                }
                else if (errCode == "3")
                {
                    if(errMessage == "전송 성공")
                    {
                        sQry = "Update [@PS_PP756H] set U_ChkYN ='Y' where DocEntry ='" + oDS_PS_PP756H.GetValue("DocEntry", 0).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);

                        sQry = "Update [@PS_PP756L] set U_TransYN ='Y' where DocEntry ='" + oDS_PS_PP756H.GetValue("DocEntry", 0).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                //    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
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
                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
            string errMessage = string.Empty;
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Button01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP756_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            PS_PP756_MTX01();
                        }
                    }
                    else if (pVal.ItemUID == "Button02")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (oDS_PS_PP756L.Size < 1)
                            {
                                errMessage = "조회 누르르고 추가하세오!";
                                BubbleEvent = false;
                                throw new System.Exception();
                            }
                            PS_PP756_InterfaceB1toR3();
                        }
                    }
                    else if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP756_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            oMat01.LoadFromDataSource();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP756_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_PP756_EnableFormItem();
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                PSH_Globals.SBO_Application.ActivateMenuItem("1291");
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_PP756_EnableFormItem();
                            }
                        }
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
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
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
                    PS_PP756_EnableFormItem();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP756H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP756L);
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
                        case "1287": //복제
                            
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
                            PS_PP756_EnableFormItem();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_PP756_EnableFormItem();
                           //PS_PP756_AddMatrixRow(0, true);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_PP756_EnableFormItem();
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
