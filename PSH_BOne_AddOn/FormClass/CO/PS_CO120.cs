using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 임금피크 대상자 현황
    /// </summary>
    internal class PS_CO120 : PSH_BaseClass
    {
        public string oFormUniqueID01;
        public SAPbouiCOM.Form oForm01;
        public SAPbouiCOM.Matrix oMat01;
        //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_CO120H;
        //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_CO120L;

        //클래스에서 선택한 마지막 아이템 Uid값
        private string oLast_Item_UID;
        //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private string oLast_Col_UID;
        //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private int oLast_Col_Row;

        private int oLast_Mode;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFromDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO120.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PS_CO120_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PS_CO120");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                //************************************************************************************************************
                //화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
                oForm01.DataBrowser.BrowseBy = "Code";
                //************************************************************************************************************


                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                PS_CO120_CreateItems();
                PS_CO120_ComboBox_Setting();
                //FormItemEnabled();

                oForm01.EnableMenu(("1283"), true);               //// 삭제
                oForm01.EnableMenu(("1287"), true);               //// 복제
                oForm01.EnableMenu(("1286"), false);              //// 닫기
                oForm01.EnableMenu(("1284"), false);              //// 취소
                oForm01.EnableMenu(("1293"), true);               //// 행삭제
                oForm01.Update();
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
                oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PS_CO120_CreateItems()
        {
            try
            {
                oForm.Freeze(true);
                ////디비데이터 소스 개체 할당
                oDS_PS_CO120H = oForm01.DataSources.DBDataSources.Item("@PS_CO120H");
                oDS_PS_CO120L = oForm01.DataSources.DBDataSources.Item("@PS_CO120L");

                //// 메트릭스 개체 할당
                oMat01 = oForm01.Items.Item("Mat01").Specific;

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
        private void PS_CO120_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                }
                else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                }
                else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
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
        /// 콤보박스 Setting
        /// </summary>
        private void PS_CO120_ComboBox_Setting()
        {
            try
            {
                SAPbouiCOM.ComboBox oCombo = null;
                string sQry = String.Empty;

                SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                //// 마감년월
                //    Set oCombo = oForm01.Items("ClsPrd").Specific
                //    sQry = "SELECT Code, Name From [OFPR]"
                //    oRecordSet01.DoQuery sQry
                //    Do Until oRecordSet01.EOF
                //        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
                //        oRecordSet01.MoveNext
                //    Loop

                //// 사업장
                oCombo = oForm01.Items.Item("BPLId").Specific;
                sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    oCombo.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
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

        public void LoadData()
        {
            try
            {
                short i = 0;
                string sQry = null;
                SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string YM = null;
                string BPLID = null;

                YM = oForm01.Items.Item("YM").Specific.VALUE.ToString().Trim();
                BPLID = oForm01.Items.Item("BPLId").Specific.VALUE.ToString().Trim();

                oForm01.Freeze(true);
                SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

                sQry = "EXEC [PS_CO120_01] '" + YM + "','" + BPLID + "'";
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_CO120L.Clear();

                if ((oRecordSet01.RecordCount == 0))
                {
                    PSH_Globals.SBO_Application.MessageBox("조회 결과가 없습니다. 확인하세요.");
                    oRecordSet01 = null;
                    oForm01.Freeze(false);
                    return;

                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_CO120L.Size)
                    {
                        oDS_PS_CO120L.InsertRecord((i));
                    }

                    oMat01.AddRow();
                    oDS_PS_CO120L.Offset = i;
                    oDS_PS_CO120L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_CO120L.SetValue("U_POEntry", i, oRecordSet01.Fields.Item("POEntry").Value.ToString().Trim());
                    oDS_PS_CO120L.SetValue("U_POLine", i, oRecordSet01.Fields.Item("POLine").Value).ToString().Trim();
                    oDS_PS_CO120L.SetValue("U_Sequence", i, oRecordSet01.Fields.Item("Sequence").Value.ToString().Trim());
                    oDS_PS_CO120L.SetValue("U_ItemCode", i, oRecordSet01.Fields.Item("ItemCode").Value.ToString().Trim());
                    oDS_PS_CO120L.SetValue("U_ItemName", i, oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim());
                    oDS_PS_CO120L.SetValue("U_CpCode", i, oRecordSet01.Fields.Item("CpCode").Value.ToString().Trim());
                    oDS_PS_CO120L.SetValue("U_CpName", i, oRecordSet01.Fields.Item("CpName").Value.ToString().Trim());
                    oDS_PS_CO120L.SetValue("U_CCCode", i, oRecordSet01.Fields.Item("CCCode").Value.ToString().Trim());
                    oDS_PS_CO120L.SetValue("U_CCName", i, oRecordSet01.Fields.Item("CCName").Value.ToString().Trim());
                    oDS_PS_CO120L.SetValue("U_ProdQty", i, oRecordSet01.Fields.Item("ProdQty").Value.ToString().Trim());
                    oDS_PS_CO120L.SetValue("U_DefQty", i, oRecordSet01.Fields.Item("DefQty").Value.ToString().Trim());
                    oDS_PS_CO120L.SetValue("U_Cost", i, oRecordSet01.Fields.Item("Cost").Value.ToString().Trim());
                    oDS_PS_CO120L.SetValue("U_Scrap", i, oRecordSet01.Fields.Item("Scrap").Value.ToString().Trim());
                    oDS_PS_CO120L.SetValue("U_Loss", i, oRecordSet01.Fields.Item("Loss").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                ProgBar01.Stop();
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

        public void FindForm()
        {
            string BPLID = string.Empty;
            string StdYM = string.Empty;
            try
            {
                //******************************************************************************
                //Function ID : PS_CO120
                //해당모듈    : FindForm
                //기능        : 제품별원가계산
                //인수        : 없음
                //반환값      : 없음
                //특이사항    : 없음
                //******************************************************************************

                oForm01.Freeze(true);
                BPLID = oForm01.Items.Item("BPLId").Specific.VALUE.ToString().Trim();
                StdYM = oForm01.Items.Item("YM").Specific.VALUE;

                //찾기모드 변경
                oForm01.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                oForm01.Items.Item("BPLId").Specific.Select(BPLID);
                oForm01.Items.Item("YM").Specific.VALUE = StdYM;

                oForm01.Items.Item("1").Click();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("FindForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public void SaveData()
        {
            try
            {
                //******************************************************************************
                //Function ID : PS_CO120
                //해당모듈    : SaveData
                //기능        : 제품별원가계산 결과 저장
                //인수        : 없음
                //반환값      : 없음
                //특이사항    : 없음
                //******************************************************************************
                // ERROR: Not supported in C#: OnErrorStatement


                oForm01.Freeze(true);

                int i = 0;
                string sQry = null;
                SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string YM = null;
                string BPLID = null;
                string UserSign = null;

                SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("저장 중...", oRecordSet01.RecordCount, false);
                YM = oForm01.Items.Item("YM").Specific.VALUE.ToString().Trim();
                BPLID = oForm01.Items.Item("BPLId").Specific.VALUE.ToString().Trim();
                UserSign = PSH_Globals.oCompany.UserSignature.ToString();

                sQry = "      EXEC [PS_CO120_50] '";
                sQry = sQry + YM + "','";
                sQry = sQry + BPLID + "','";
                sQry = sQry + UserSign + "'";

                oRecordSet01.DoQuery(sQry);
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
        /// 메트릭스 Row 추가
        /// </summary>
        public void Add_MatrixRow(int oRow, bool RowIserted = false)
        {
            try
            {
                oForm.Freeze(true);
                ////행추가여부
                if (RowIserted == false)
                {
                    oDS_PS_CO120L.InsertRecord((oRow));
                }
                oMat01.AddRow();
                oDS_PS_CO120L.Offset = oRow;
                oDS_PS_CO120L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

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

                    //case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    //    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //    Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
                int i = 0;
                if ((pVal.BeforeAction == true))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284":
                            //취소
                            break;
                        case "1286":
                            //닫기
                            break;
                        case "1293":
                            //행삭제
                            break;
                        case "1281":
                            //찾기
                            break;
                        case "1282":
                            //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            //레코드이동버튼
                            break;
                    }
                    ////BeforeAction = False
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284":
                            //취소
                            break;
                        case "1286":
                            //닫기
                            break;
                        case "1293":
                            //행삭제
                            oForm01.Freeze(true);
                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.VALUE = i + 1;
                                }

                                oMat01.FlushToDataSource();
                                oDS_PS_CO120L.RemoveRecord(oDS_PS_CO120L.Size - 1);
                                //// Mat01에 마지막라인(빈라인) 삭제
                                oMat01.Clear();
                                oMat01.LoadFromDataSource();

                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(oMat01.RowCount).Specific.VALUE))
                                {
                                    Add_MatrixRow(oMat01.RowCount, false);
                                }
                            }
                            oForm01.Freeze(false);
                            break;
                        case "1281":
                            //찾기
                            oForm01.Freeze(true);
                            PS_CO120_FormItemEnabled();
                            //                oForm01.Items("CycleCod").Click ct_Regular
                            oForm01.Freeze(false);
                            break;
                        case "1282":
                            //추가
                            oForm01.Freeze(true);
                            PS_CO120_FormItemEnabled();
                            Add_MatrixRow(0, true);
                            oForm01.Freeze(false);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            //레코드이동버튼
                            oForm01.Freeze(true);
                            PS_CO120_FormItemEnabled();
                            oForm01.Freeze(false);
                            break;
                        case "1287":
                            //// 복제
                            oForm01.Freeze(true);
                            oDS_PS_CO120H.SetValue("Code", 0, "");
                            oDS_PS_CO120H.SetValue("Name", 0, "");
                            oDS_PS_CO120H.SetValue("U_YM", 0, "");
                            oDS_PS_CO120H.SetValue("U_BPLId", 0, "");

                            for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                oMat01.FlushToDataSource();
                                oDS_PS_CO120L.SetValue("Code", i, "");
                                oMat01.LoadFromDataSource();
                            }

                            oForm01.Freeze(false);
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
            int i = 0;
            string sQry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                            ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                            ////36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                            ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                            ////36
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pVal.Action_Success == true)
                        {
                            oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PSH_Globals.SBO_Application.ActivateMenuItem("1282");
                        }
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        if (PS_CO120_HeaderSpaceLineDel() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        SaveData();                    //백그라운드(쿼리)에서 저장하는 로직으로 수정(2018.07.05 송명규)
                        FindForm();                    //계산 실행 후 결과 확인을 위한 Find Mode 변경

                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pVal.Action_Success == true)
                        {
                            oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PSH_Globals.SBO_Application.ActivateMenuItem("1282");
                        }
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        if (PS_CO120_HeaderSpaceLineDel() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        //Call LoadData
                        SaveData();                        //백그라운드(쿼리)에서 저장하는 로직으로 수정(2018.07.05 송명규)
                        FindForm();                        //계산 실행 후 결과 확인을 위한 Find Mode 변경

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
        /// DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.Row == 0)
                    {

                        oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;     //정렬
                        oMat01.FlushToDataSource();
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
        /// 필수입력사항 체크
        /// </summary>
        /// <returns>True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음</returns>
        private bool PS_CO120_HeaderSpaceLineDel()
        {
            bool functionReturnValue = false;
            short ErrNum = 0;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_CO120H.GetValue("U_YM", 0).ToString().Trim())){
                    ErrNum = 1;
                    throw new Exception();
                }
                        
                if (string.IsNullOrEmpty(oDS_PS_CO120H.GetValue("U_BPLId", 0).ToString().Trim())){
                    ErrNum = 2;
                    throw new Exception();
                }
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("마감년월은 필수입력사항입니다. 확인하세요.");
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장은 필수입력사항입니다. 확인하세요.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                functionReturnValue = false;
            }
            finally
            {
            }

            return functionReturnValue;
        }
    }
    }

//using System.Collections.Generic;
//using SAPbouiCOM;
//using PSH_BOne_AddOn.Data;
//using PSH_BOne_AddOn.DataPack;
//using PSH_BOne_AddOn.Form;

//namespace PSH_BOne_AddOn
//{
//    /// <summary>
//    /// 임금피크 대상자 현황
//    /// </summary>
//    internal class PS_CO120 : PSH_BaseClass
//    {
//        public string oFormUniqueID01;
//        public SAPbouiCOM.Form oForm01;
//        public SAPbouiCOM.Matrix oMat01;
//        //등록헤더
//        private SAPbouiCOM.DBDataSource oDS_PS_CO120H;
//        //등록라인
//        private SAPbouiCOM.DBDataSource oDS_PS_CO120L;

//        //클래스에서 선택한 마지막 아이템 Uid값
//        private string oLast_Item_UID;
//        //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
//        private string oLast_Col_UID;
//        //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
//        private int oLast_Col_Row;

//        private int oLast_Mode;

//        /// <summary>
//        /// 화면 호출
//        /// </summary>
//        public override void LoadForm(string oFromDocEntry01)
//        {
//            int i = 0;
//            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
//            try
//            {
//                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO120.srf");
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

//                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
//                {
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//                }

//                oFormUniqueID01 = "PS_CO120_" + SubMain.Get_TotalFormsCount();
//                SubMain.Add_Forms(this, oFormUniqueID01, "PS_CO120");

//                string strXml = string.Empty;
//                strXml = oXmlDoc.xml.ToString();

//                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
//                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

//                //************************************************************************************************************
//                //화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
//                oForm01.DataBrowser.BrowseBy = "Code";
//                //************************************************************************************************************


//                oForm.SupportedModes = -1;
//                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

//                PS_CO120_CreateItems();
//                ComboBox_Setting();
//                //FormItemEnabled();

//                oForm01.EnableMenu(("1283"), true);               //// 삭제
//                oForm01.EnableMenu(("1287"), true);               //// 복제
//                oForm01.EnableMenu(("1286"), false);              //// 닫기
//                oForm01.EnableMenu(("1284"), false);              //// 취소
//                oForm01.EnableMenu(("1293"), true);               //// 행삭제
//                oForm01.Update();
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Update();
//                oForm.Freeze(false);
//                oForm.Visible = true;
//                oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
//            }
//        }

//        /// <summary>
//        /// 화면 Item 생성
//        /// </summary>
//        /// <returns></returns>
//        private void PS_CO120_CreateItems()
//        {
//            try
//            {
//                oForm.Freeze(true);
//                ////디비데이터 소스 개체 할당
//                oDS_PS_CO120H = oForm01.DataSources.DBDataSources.Item("@PS_CO120H");
//                oDS_PS_CO120L = oForm01.DataSources.DBDataSources.Item("@PS_CO120L");

//                //// 메트릭스 개체 할당
//                oMat01 = oForm01.Items.Item("Mat01").Specific;

//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY522_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        /// <summary>
//        /// 화면의 아이템 Enable 설정
//        /// </summary>
//        private void PS_CO120_FormItemEnabled()
//        {
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                oForm.Freeze(true);
//                if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
//                {
//                }
//                else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
//                {
//                }
//                else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
//                {
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY003_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }


//        /// <summary>
//        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
//        /// </summary>
//        /// <param name="oUID"></param>
//        /// <param name="oRow"></param>
//        /// <param name="oCol"></param>
//        private void PH_PY011_FlushToItemValue(string oUID, int oRow, string oCol)
//        {
//            int i = 0;
//            short ErrNum = 0;
//            string sQry = null;
//            int sRow = 0;
//            string sSeq = null;
//            sRow = oRow;

//            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                switch (oUID)
//                {
//                    case "Mat01":
//                        break;
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY011_FlushToItemValue_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
//            }
//        }

//        public void LoadData()
//        {
//            try
//            {
//                short i = 0;
//                string sQry = null;
//                SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//                string YM = null;
//                string BPLID = null;

//                YM = oForm01.Items.Item("YM").Specific.VALUE.ToString().Trim();
//                BPLID = oForm01.Items.Item("BPLId").Specific.VALUE.ToString().Trim();

//                oForm01.Freeze(true);
//                SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

//                sQry = "EXEC [PS_CO120_01] '" + YM + "','" + BPLID + "'";
//                oRecordSet01.DoQuery(sQry);

//                oMat01.Clear();
//                oDS_PS_CO120L.Clear();

//                if ((oRecordSet01.RecordCount == 0))
//                {
//                    MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.:" + Err().Number + " - " + Err().Description, ref "W");
//                    oRecordSet01 = null;
//                    oForm01.Freeze(false);
//                    return;

//                }

//                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
//                {
//                    if (i + 1 > oDS_PS_CO120L.Size)
//                    {
//                        oDS_PS_CO120L.InsertRecord((i));
//                    }

//                    oMat01.AddRow();
//                    oDS_PS_CO120L.Offset = i;
//                    oDS_PS_CO120L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//                    oDS_PS_CO120L.SetValue("U_POEntry", i, Strings.Trim(oRecordSet01.Fields.Item("POEntry").Value));
//                    oDS_PS_CO120L.SetValue("U_POLine", i, Strings.Trim(oRecordSet01.Fields.Item("POLine").Value));
//                    oDS_PS_CO120L.SetValue("U_Sequence", i, Strings.Trim(oRecordSet01.Fields.Item("Sequence").Value));
//                    oDS_PS_CO120L.SetValue("U_ItemCode", i, Strings.Trim(oRecordSet01.Fields.Item("ItemCode").Value));
//                    oDS_PS_CO120L.SetValue("U_ItemName", i, Strings.Trim(oRecordSet01.Fields.Item("ItemName").Value));
//                    oDS_PS_CO120L.SetValue("U_CpCode", i, Strings.Trim(oRecordSet01.Fields.Item("CpCode").Value));
//                    oDS_PS_CO120L.SetValue("U_CpName", i, Strings.Trim(oRecordSet01.Fields.Item("CpName").Value));
//                    oDS_PS_CO120L.SetValue("U_CCCode", i, Strings.Trim(oRecordSet01.Fields.Item("CCCode").Value));
//                    oDS_PS_CO120L.SetValue("U_CCName", i, Strings.Trim(oRecordSet01.Fields.Item("CCName").Value));
//                    oDS_PS_CO120L.SetValue("U_ProdQty", i, Strings.Trim(oRecordSet01.Fields.Item("ProdQty").Value));
//                    oDS_PS_CO120L.SetValue("U_DefQty", i, Strings.Trim(oRecordSet01.Fields.Item("DefQty").Value));
//                    oDS_PS_CO120L.SetValue("U_Cost", i, Strings.Trim(oRecordSet01.Fields.Item("Cost").Value));
//                    oDS_PS_CO120L.SetValue("U_Scrap", i, Strings.Trim(oRecordSet01.Fields.Item("Scrap").Value));
//                    oDS_PS_CO120L.SetValue("U_Loss", i, Strings.Trim(oRecordSet01.Fields.Item("Loss").Value));

//                    oRecordSet01.MoveNext();
//                    ProgBar01.Value = ProgBar01.Value + 1;
//                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
//                }
//                oMat01.LoadFromDataSource();
//                oMat01.AutoResizeColumns();
//                ProgBar01.Stop();
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY002_AddMatrixRow_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }


//        public void SaveData()
//        {
//            try
//            {
//                //******************************************************************************
//                //Function ID : PS_CO120
//                //해당모듈    : SaveData
//                //기능        : 제품별원가계산 결과 저장
//                //인수        : 없음
//                //반환값      : 없음
//                //특이사항    : 없음
//                //******************************************************************************
//                // ERROR: Not supported in C#: OnErrorStatement


//                oForm01.Freeze(true);

//                int i = 0;
//                string sQry = null;

//                SAPbobsCOM.Recordset oRecordSet01 = null;
//                oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//                string YM = null;
//                string BPLID = null;
//                string UserSign = null;

//                SAPbouiCOM.ProgressBar ProgBar01 = null;
//                ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("저장 중...", 100, false);

//                //UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                YM = Strings.Trim(oForm01.Items.Item("YM").Specific.VALUE);
//                //UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                BPLID = Strings.Trim(oForm01.Items.Item("BPLId").Specific.VALUE);
//                UserSign = Convert.ToString(SubMain.Sbo_Company.UserSignature);

//                sQry = "      EXEC [PS_CO120_50] '";
//                sQry = sQry + YM + "','";
//                sQry = sQry + BPLID + "','";
//                sQry = sQry + UserSign + "'";

//                oRecordSet01.DoQuery(sQry);
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY002_AddMatrixRow_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        /// <summary>
//        /// 메트릭스 Row 추가
//        /// </summary>
//        public void Add_MatrixRow(int oRow, bool RowIserted = false)
//        {
//            try
//            {
//                oForm.Freeze(true);
//                ////행추가여부
//                if (RowIserted == false)
//                {
//                    oDS_PS_CO120L.InsertRecord((oRow));
//                }
//                oMat01.AddRow();
//                oDS_PS_CO120L.Offset = oRow;
//                oDS_PS_CO120L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//                oMat01.LoadFromDataSource();
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY002_AddMatrixRow_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        /// <summary>
//        /// Form Item Event
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">pVal</param>
//        /// <param name="BubbleEvent">Bubble Event</param>
//        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            switch (pVal.EventType)
//            {
//                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
//                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                    //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
//                    //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
//                    //    break;

//                    //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
//                    //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
//                    //    break;

//                    //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
//                    //    break;

//                    //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
//                    //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
//                    //    break;

//                    //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
//                    //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
//                    //    break;

//                    //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
//                    //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
//                    //    break;

//                    //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
//                    //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
//                    //    break;

//                    ////case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
//                    ////    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
//                    ////    break;

//                    //case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
//                    //    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
//                    //    break;

//                    //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
//                    //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
//                    //    break;

//                    ////case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
//                    ////    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
//                    ////    break;

//                    //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
//                    //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
//                    //    break;

//                    //case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
//                    //    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
//                    //    break;

//                    //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
//                    //    break;

//                    //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
//                    //    break;

//                    ////case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
//                    ////    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
//                    ////    break;

//                    //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
//                    //    Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
//                    //    break;

//                    ////case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
//                    ////    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
//                    ////    break;

//                    ////case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
//                    ////    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
//                    ////    break;

//                    //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
//                    //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
//                    //    break;

//                    //    //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
//                    //    //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
//                    //    //    break;

//                    //    //case SAPbouiCOM.BoEventTypes.et_Drag: //39
//                    //    //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
//                    //    //    break;
//            }
//        }

//        /// <summary>
//        /// FormMenuEvent
//        /// </summary>
//        /// <param name="FormUID"></param>
//        /// <param name="pVal"></param>
//        /// <param name="BubbleEvent"></param>
//        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
//        {
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                oForm.Freeze(true);
//                int i = 0;
//                if ((pVal.BeforeAction == true))
//                {
//                    switch (pVal.MenuUID)
//                    {
//                        case "1284":
//                            //취소
//                            break;
//                        case "1286":
//                            //닫기
//                            break;
//                        case "1293":
//                            //행삭제
//                            break;
//                        case "1281":
//                            //찾기
//                            break;
//                        case "1282":
//                            //추가
//                            break;
//                        case "1288":
//                        case "1289":
//                        case "1290":
//                        case "1291":
//                            //레코드이동버튼
//                            break;
//                    }
//                    ////BeforeAction = False
//                }
//                else if ((pVal.BeforeAction == false))
//                {
//                    switch (pVal.MenuUID)
//                    {
//                        case "1284":
//                            //취소
//                            break;
//                        case "1286":
//                            //닫기
//                            break;
//                        case "1293":
//                            //행삭제
//                            oForm01.Freeze(true);
//                            if (oMat01.RowCount != oMat01.VisualRowCount)
//                            {
//                                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
//                                {
//                                    oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.VALUE = i + 1;
//                                }

//                                oMat01.FlushToDataSource();
//                                oDS_PS_CO120L.RemoveRecord(oDS_PS_CO120L.Size - 1);
//                                //// Mat01에 마지막라인(빈라인) 삭제
//                                oMat01.Clear();
//                                oMat01.LoadFromDataSource();

//                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(oMat01.RowCount).Specific.VALUE))
//                                {
//                                    Add_MatrixRow(oMat01.RowCount, ref false);
//                                }
//                            }
//                            oForm01.Freeze(false);
//                            break;
//                        case "1281":
//                            //찾기
//                            oForm01.Freeze(true);
//                            FormItemEnabled();
//                            //                oForm01.Items("CycleCod").Click ct_Regular
//                            oForm01.Freeze(false);
//                            break;
//                        case "1282":
//                            //추가
//                            oForm01.Freeze(true);
//                            FormItemEnabled();
//                            Add_MatrixRow(0, ref true);
//                            oForm01.Freeze(false);
//                            break;
//                        case "1288":
//                        case "1289":
//                        case "1290":
//                        case "1291":
//                            //레코드이동버튼
//                            oForm01.Freeze(true);
//                            FormItemEnabled();
//                            oForm01.Freeze(false);
//                            break;
//                        case "1287":
//                            //// 복제
//                            oForm01.Freeze(true);
//                            oDS_PS_CO120H.SetValue("Code", 0, "");
//                            oDS_PS_CO120H.SetValue("Name", 0, "");
//                            oDS_PS_CO120H.SetValue("U_YM", 0, "");
//                            oDS_PS_CO120H.SetValue("U_BPLId", 0, "");

//                            for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
//                            {
//                                oMat01.FlushToDataSource();
//                                oDS_PS_CO120L.SetValue("Code", i, "");
//                                oMat01.LoadFromDataSource();
//                            }

//                            oForm01.Freeze(false);
//                            break;
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormMenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }


//        /// <summary>
//        /// FormDataEvent
//        /// </summary>
//        /// <param name="FormUID"></param>
//        /// <param name="BusinessObjectInfo"></param>
//        /// <param name="BubbleEvent"></param>
//        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//        {
//            int i = 0;
//            string sQry = string.Empty;

//            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                if (BusinessObjectInfo.BeforeAction == true)
//                {
//                    switch (BusinessObjectInfo.EventType)
//                    {
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            ////33
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                            ////34
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                            ////35
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                            ////36
//                            break;
//                    }
//                }
//                else if (BusinessObjectInfo.BeforeAction == false)
//                {
//                    switch (BusinessObjectInfo.EventType)
//                    {
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            ////33
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                            ////34
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                            ////35
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                            ////36
//                            break;
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
//            }
//        }

//        /// <summary>
//        /// 콤보박스 Setting
//        /// </summary>
//        private void PS_CO120_ComboBox_Setting()
//        {
//            try
//            {
//                oForm.Freeze(true);
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY011_ComboBox_Setting_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        /// <summary>
//        /// ITEM_PRESSED 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {

//                if (pVal.Before_Action == true)
//                {
//                    if (pVal.ItemUID == "1")
//                    {
//                        if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pVal.Action_Success == true)
//                        {
//                            oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//                            SubMain.Sbo_Application.ActivateMenuItem("1282");
//                        }
//                    }
//                    else if (pVal.ItemUID == "Btn01")
//                    {
//                        if (HeaderSpaceLineDel() == false)
//                        {
//                            BubbleEvent = false;
//                            return;
//                        }
//                        SaveData();                    //백그라운드(쿼리)에서 저장하는 로직으로 수정(2018.07.05 송명규)
//                        FindForm();                    //계산 실행 후 결과 확인을 위한 Find Mode 변경

//                    }
//                }
//                else if (pVal.Before_Action == false)
//                {
//                    if (pVal.ItemUID == "1")
//                    {
//                        if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pVal.Action_Success == true)
//                        {
//                            oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//                            SubMain.Sbo_Application.ActivateMenuItem("1282");
//                        }
//                    }
//                    else if (pVal.ItemUID == "Btn01")
//                    {
//                        if (HeaderSpaceLineDel() == false)
//                        {
//                            BubbleEvent = false;
//                            return;
//                        }
//                        //Call LoadData
//                        SaveData();                        //백그라운드(쿼리)에서 저장하는 로직으로 수정(2018.07.05 송명규)
//                        FindForm();                        //계산 실행 후 결과 확인을 위한 Find Mode 변경

//                    }
//                }

//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// DOUBLE_CLICK 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.Before_Action == true)
//                {
//                    if (pVal.Row == 0)
//                    {

//                        oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;     //정렬
//                        oMat01.FlushToDataSource();
//                    }
//                }
//                else if (pVal.Before_Action == false)
//                {
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_DOUBLE_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// 필수입력사항 체크
//        /// </summary>
//        /// <returns>True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음</returns>
//        private bool PH_PY011_HeaderSpaceLineDel()
//        {
//            bool functionReturnValue = false;
//            short ErrNum = 0;

//            try
//            {
//                switch (true)
//                {
//                    case string.IsNullOrEmpty(oDS_PS_CO120H.GetValue("U_YM", 0)):
//                        ErrNum = 1;
//                        goto HeaderSpaceLineDel_Error;
//                        break;
//                    case string.IsNullOrEmpty(oDS_PS_CO120H.GetValue("U_BPLId", 0)):
//                        ErrNum = 2;
//                        goto HeaderSpaceLineDel_Error;
//                        break;
//                }
//                functionReturnValue = true;
//            }
//            catch (Exception ex)
//            {
//                if (ErrNum == 1)
//                {
//                    MDC_Com.MDC_GF_Message(ref "마감년월은 필수입력사항입니다. 확인하세요.", ref "E");
//                }
//                else if (ErrNum == 2)
//                {
//                    MDC_Com.MDC_GF_Message(ref "사업장은 필수입력사항입니다. 확인하세요.", ref "E");
//                }
//                else
//                {
//                    MDC_Com.MDC_GF_Message(ref "HeaderSpaceLineDel_Error:" + Err().Number + " - " + Err().Description, ref "E");
//                }
//                functionReturnValue = false;
//            }
//            finally
//            {
//            }

//            return functionReturnValue;
//        }

//        /// <summary>
//        /// 메트릭스 필수 사항 check
//        /// 구현은 되어 있지만 사용하지 않음
//        /// </summary>
//        /// <returns></returns>
//        private bool PH_PY011_MatrixSpaceLineDel()
//        {
//            bool functionReturnValue = false;

//            int i = 0;
//            short ErrNum = 0;
//            SAPbobsCOM.Recordset oRecordSet = null;
//            string sQry = null;

//            try
//            {
//                oMat01.FlushToDataSource();

//                //// 라인
//                if (oMat01.VisualRowCount == 0)
//                {
//                    ErrNum = 1;
//                    goto MatrixSpaceLineDel_Error;
//                }
//                else if (oMat01.VisualRowCount == 1)
//                {
 
//                }

//                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
//                {

//                }
//                oMat01.LoadFromDataSource();
//                functionReturnValue = true;
//            }
//            catch (Exception ex)
//            {
//                if (ErrNum == 1)
//                {
//                    MDC_Com.MDC_GF_Message(ref "라인 데이터가 없습니다. 확인하세요.", ref "E");
//                }
//                else if (ErrNum == 2)
//                {
//                    MDC_Com.MDC_GF_Message(ref "첫라인에 배부사이클 코드가 없습니다. 확인하세요.", ref "E");
//                }
//                else if (ErrNum == 3)
//                {
//                    MDC_Com.MDC_GF_Message(ref "수량은 필수사항입니다. 확인하세요.", ref "E");
//                }
//                else if (ErrNum == 4)
//                {
//                    MDC_Com.MDC_GF_Message(ref "중량은 필수사항입니다. 확인하세요.", ref "E");
//                }
//                else if (ErrNum == 5)
//                {
//                    MDC_Com.MDC_GF_Message(ref "단가는 필수사항입니다. 확인하세요.", ref "E");
//                }
//                else if (ErrNum == 6)
//                {
//                    MDC_Com.MDC_GF_Message(ref "금액은 필수사항입니다. 확인하세요.", ref "E");
//                }
//                else
//                {
//                    MDC_Com.MDC_GF_Message(ref "MatrixSpaceLineDel_Error:" + Err().Number + " - " + Err().Description, ref "E");
//                }

//                functionReturnValue = false;
//            }

//            return functionReturnValue;
//        }


//    }
//}



////        using Microsoft.VisualBasic;
////using Microsoft.VisualBasic.Compatibility;
////using System;
////using System.Collections;
////using System.Data;
////using System.Diagnostics;
////using System.Drawing;
////using System.Windows.Forms;
//// // ERROR: Not supported in C#: OptionDeclaration
////namespace MDC_PS_Addon
////	{
////		internal class PS_CO120
////		{
////			//****************************************************************************************************************
////			////  File           : PS_CO120.cls
////			////  Module         : CO
////			////  Description    : 공정별 원가계산
////			////  FormType       : PS_CO120
////			////  Create Date    : 2010.11.17
////			////  Modified Date  :
////			////  Creator        : Ryu Yung Jo
////			////  Company        : Poongsan Holdings
////			//****************************************************************************************************************

////			public string oFormUniqueID01;
////			public SAPbouiCOM.Form oForm01;
////			public SAPbouiCOM.Matrix oMat01;
////			//등록헤더
////			private SAPbouiCOM.DBDataSource oDS_PS_CO120H;
////			//등록라인
////			private SAPbouiCOM.DBDataSource oDS_PS_CO120L;

////			//클래스에서 선택한 마지막 아이템 Uid값
////			private string oLast_Item_UID;
////			//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
////			private string oLast_Col_UID;
////			//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
////			private int oLast_Col_Row;

////			private int oLast_Mode;

////			//****************************************************************************************************************
////			// .srf 파일로부터 폼을 로드한다.
////			//****************************************************************************************************************
////			public void LoadForm()
////			{
////				// ERROR: Not supported in C#: OnErrorStatement

////				int i = 0;
////				string oInnerXml01 = null;
////				MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

////				oXmlDoc01.load(SubMain.ShareFolderPath + "ScreenPS\\PS_CO120.srf");
////				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
////				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
////				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);

////				//매트릭스의 타이틀높이와 셀높이를 고정
////				for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
////				{
////					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
////					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
////				}

////				oFormUniqueID01 = "PS_CO120_" + GetTotalFormsCount();
////				SubMain.AddForms(this, oFormUniqueID01);
////				////폼추가
////				SubMain.Sbo_Application.LoadBatchActions(out (oXmlDoc01.xml));

////				//폼 할당
////				oForm01 = SubMain.Sbo_Application.Forms.Item(oFormUniqueID01);

////				oForm01.SupportedModes = -1;
////				oForm01.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				//************************************************************************************************************
////				//화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
////				oForm01.DataBrowser.BrowseBy = "Code";
////				//************************************************************************************************************
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////

////				oForm01.Freeze(true);
////				CreateItems();
////				ComboBox_Setting();
////				//    Call Add_MatrixRow(0, True)
////				FormItemEnabled();

////				oForm01.EnableMenu(("1283"), true);
////				//// 삭제
////				oForm01.EnableMenu(("1287"), true);
////				//// 복제
////				oForm01.EnableMenu(("1286"), false);
////				//// 닫기
////				oForm01.EnableMenu(("1284"), false);
////				//// 취소
////				oForm01.EnableMenu(("1293"), true);
////				//// 행삭제

////				oForm01.Update();

////				oForm01.Freeze(false);
////				oForm01.Visible = true;
////				//UPGRADE_NOTE: oXmlDoc01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				oXmlDoc01 = null;
////				return;
////			LoadForm_Error:
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				oForm01.Update();
////				oForm01.Freeze(false);
////				//UPGRADE_NOTE: oXmlDoc01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				oXmlDoc01 = null;
////				if ((oForm01 == null) == false)
////				{
////					//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////					oForm01 = null;
////				}
////				MDC_Com.MDC_GF_Message(ref "LoadForm_Error:" + Err().Number + " - " + Err().Description, ref "E");
////			}

////			//****************************************************************************************************************
////			//// ItemEventHander
////			//****************************************************************************************************************
////			public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
////			{
////				// ERROR: Not supported in C#: OnErrorStatement

////				int i = 0;
////				string ItemType = null;
////				string RequestDate = null;
////				string Size = null;
////				string ItemCode = null;
////				string ItemName = null;
////				string Unit = null;
////				string DueDate = null;
////				string RequestNo = null;
////				int Qty = 0;
////				decimal Weight = default(decimal);
////				double Calculate_Weight = 0;

////				object ChildForm01 = null;
////				ChildForm01 = new PS_CO111();

////				string BPLID = null;
////				string YM = null;
////				string Code = null;
////				////BeforeAction = True
////				if ((pVal.BeforeAction == true))
////				{
////					switch (pVal.EventType)
////					{
////						//et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
////						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
////							////1
////							if (pVal.ItemUID == "1")
////							{
////								if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
////								{
////									if (HeaderSpaceLineDel() == false)
////									{
////										BubbleEvent = false;
////										return;
////									}
////									if (MatrixSpaceLineDel() == false)
////									{
////										BubbleEvent = false;
////										return;
////									}
////									YM = Strings.Trim(oDS_PS_CO120H.GetValue("U_YM", 0));
////									BPLID = Strings.Trim(oDS_PS_CO120H.GetValue("U_BPLId", 0));
////									Code = YM + BPLID;
////									oDS_PS_CO120H.SetValue("Code", 0, Code);
////									oDS_PS_CO120H.SetValue("Name", 0, Code);
////								}
////							}
////							break;
////						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
////							////2
////							break;
////						case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
////							////5
////							break;
////						case SAPbouiCOM.BoEventTypes.et_CLICK:
////							////6
////							break;
////						case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
////							////7

////							if (pVal.Row == 0)
////							{

////								//정렬
////								oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
////								oMat01.FlushToDataSource();

////							}
////							break;

////						case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
////							////8
////							break;
////						case SAPbouiCOM.BoEventTypes.et_VALIDATE:
////							////10
////							break;
////						case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
////							////11
////							break;
////						case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
////							////18
////							break;
////						case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
////							////19
////							break;
////						case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
////							////20
////							break;
////						case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
////							////27
////							break;
////						case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
////							////3
////							break;
////						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
////							////4
////							break;
////						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
////							////17
////							break;
////					}
////					////BeforeAction = False
////				}
////				else if ((pVal.BeforeAction == false))
////				{
////					switch (pVal.EventType)
////					{
////						//et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
////						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
////							////1
////							if (pVal.ItemUID == "1")
////							{
////								if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pVal.Action_Success == true)
////								{
////									oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
////									SubMain.Sbo_Application.ActivateMenuItem("1282");
////								}
////							}
////							else if (pVal.ItemUID == "Btn01")
////							{
////								if (HeaderSpaceLineDel() == false)
////								{
////									BubbleEvent = false;
////									return;
////								}
////								//Call LoadData
////								SaveData();
////								//백그라운드(쿼리)에서 저장하는 로직으로 수정(2018.07.05 송명규)
////								FindForm();
////								//계산 실행 후 결과 확인을 위한 Find Mode 변경

////							}
////							break;
////						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
////							////2
////							break;
////						case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
////							////5
////							break;
////						case SAPbouiCOM.BoEventTypes.et_CLICK:
////							////6
////							break;
////						case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
////							////8
////							break;
////						case SAPbouiCOM.BoEventTypes.et_VALIDATE:
////							////10
////							break;
////						//et_MATRIX_LOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
////						case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
////							////11
////							Add_MatrixRow(oMat01.RowCount, ref false);
////							break;
////						case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
////							////18
////							break;
////						case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
////							////19
////							break;
////						case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
////							////20
////							break;
////						case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
////							////27
////							break;
////						case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
////							////3
////							break;
////						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
////							////4
////							break;
////						//et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
////						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
////							////17
////							SubMain.RemoveForms(oFormUniqueID01);
////							//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////							oForm01 = null;
////							//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////							oMat01 = null;
////							//UPGRADE_NOTE: oDS_PS_CO120H 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////							oDS_PS_CO120H = null;
////							//UPGRADE_NOTE: oDS_PS_CO120L 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////							oDS_PS_CO120L = null;
////							break;
////					}
////				}
////				return;
////			Raise_ItemEvent_Error:
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				oForm01.Freeze(false);
////				MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
////			}

////			public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
////			{
////				// ERROR: Not supported in C#: OnErrorStatement

////				int i = 0;

////				////BeforeAction = True
////				if ((pVal.BeforeAction == true))
////				{
////					switch (pVal.MenuUID)
////					{
////						case "1284":
////							//취소
////							break;
////						case "1286":
////							//닫기
////							break;
////						case "1293":
////							//행삭제
////							break;
////						case "1281":
////							//찾기
////							break;
////						case "1282":
////							//추가
////							break;
////						case "1288":
////						case "1289":
////						case "1290":
////						case "1291":
////							//레코드이동버튼
////							break;
////					}
////					////BeforeAction = False
////				}
////				else if ((pVal.BeforeAction == false))
////				{
////					switch (pVal.MenuUID)
////					{
////						case "1284":
////							//취소
////							break;
////						case "1286":
////							//닫기
////							break;
////						case "1293":
////							//행삭제
////							oForm01.Freeze(true);
////							if (oMat01.RowCount != oMat01.VisualRowCount)
////							{
////								for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
////								{
////									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
////									oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.VALUE = i + 1;
////								}

////								oMat01.FlushToDataSource();
////								oDS_PS_CO120L.RemoveRecord(oDS_PS_CO120L.Size - 1);
////								//// Mat01에 마지막라인(빈라인) 삭제
////								oMat01.Clear();
////								oMat01.LoadFromDataSource();

////								//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(oMat01.RowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
////								if (!string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(oMat01.RowCount).Specific.VALUE))
////								{
////									Add_MatrixRow(oMat01.RowCount, ref false);
////								}
////							}
////							oForm01.Freeze(false);
////							break;
////						case "1281":
////							//찾기
////							oForm01.Freeze(true);
////							FormItemEnabled();
////							//                oForm01.Items("CycleCod").Click ct_Regular
////							oForm01.Freeze(false);
////							break;
////						case "1282":
////							//추가
////							oForm01.Freeze(true);
////							FormItemEnabled();
////							Add_MatrixRow(0, ref true);
////							oForm01.Freeze(false);
////							break;
////						case "1288":
////						case "1289":
////						case "1290":
////						case "1291":
////							//레코드이동버튼
////							oForm01.Freeze(true);
////							FormItemEnabled();
////							//                If oMat01.VisualRowCount > 0 Then
////							//                    If oMat01.Columns("CycleCod").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
////							//                        Add_MatrixRow oMat01.RowCount, False
////							//                    End If
////							//                End If
////							oForm01.Freeze(false);
////							break;
////						case "1287":
////							//// 복제
////							oForm01.Freeze(true);
////							oDS_PS_CO120H.SetValue("Code", 0, "");
////							oDS_PS_CO120H.SetValue("Name", 0, "");
////							oDS_PS_CO120H.SetValue("U_YM", 0, "");
////							oDS_PS_CO120H.SetValue("U_BPLId", 0, "");

////							for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
////							{
////								oMat01.FlushToDataSource();
////								oDS_PS_CO120L.SetValue("Code", i, "");
////								oMat01.LoadFromDataSource();
////							}

////							oForm01.Freeze(false);
////							break;
////					}
////				}
////				return;
////			Raise_MenuEvent_Error:
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				oForm01.Freeze(false);
////				MDC_Com.MDC_GF_Message(ref "Raise_MenuEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
////			}

////			public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
////			{
////				// ERROR: Not supported in C#: OnErrorStatement

////				////BeforeAction = True
////				if ((BusinessObjectInfo.BeforeAction == true))
////				{
////					switch (BusinessObjectInfo.EventType)
////					{
////						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
////							////33
////							break;
////						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
////							////34
////							break;
////						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
////							////35
////							break;
////						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
////							////36
////							break;
////					}
////					////BeforeAction = False
////				}
////				else if ((BusinessObjectInfo.BeforeAction == false))
////				{
////					switch (BusinessObjectInfo.EventType)
////					{
////						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
////							////33
////							break;
////						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
////							////34
////							break;
////						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
////							////35
////							break;
////						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
////							////36
////							break;
////					}
////				}
////				return;
////			Raise_FormDataEvent_Error:
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				MDC_Com.MDC_GF_Message(ref "Raise_FormDataEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
////			}

////			public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
////			{
////				// ERROR: Not supported in C#: OnErrorStatement

////				if ((eventInfo.BeforeAction == true))
////				{
////					////작업
////				}
////				else if ((eventInfo.BeforeAction == false))
////				{
////					////작업
////				}
////				return;
////			Raise_RightClickEvent_Error:
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
////			}

////			private void CreateItems()
////			{
////				// ERROR: Not supported in C#: OnErrorStatement

////				////디비데이터 소스 개체 할당
////				oDS_PS_CO120H = oForm01.DataSources.DBDataSources("@PS_CO120H");
////				oDS_PS_CO120L = oForm01.DataSources.DBDataSources("@PS_CO120L");

////				//// 메트릭스 개체 할당
////				oMat01 = oForm01.Items.Item("Mat01").Specific;

////				return;
////			CreateItems_Error:
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				MDC_Com.MDC_GF_Message(ref "CreateItems_Error:" + Err().Number + " - " + Err().Description, ref "E");
////			}

////			public void ComboBox_Setting()
////			{
////				// ERROR: Not supported in C#: OnErrorStatement

////				////콤보에 기본값설정
////				SAPbouiCOM.ComboBox oCombo = null;
////				string sQry = null;
////				SAPbobsCOM.Recordset oRecordSet01 = null;

////				oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

////				//// 마감년월
////				//    Set oCombo = oForm01.Items("ClsPrd").Specific
////				//    sQry = "SELECT Code, Name From [OFPR]"
////				//    oRecordSet01.DoQuery sQry
////				//    Do Until oRecordSet01.EOF
////				//        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
////				//        oRecordSet01.MoveNext
////				//    Loop

////				//// 사업장
////				oCombo = oForm01.Items.Item("BPLId").Specific;
////				sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
////				oRecordSet01.DoQuery(sQry);
////				while (!(oRecordSet01.EoF))
////				{
////					oCombo.ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
////					oRecordSet01.MoveNext();
////				}




////				//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				oCombo = null;
////				//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				oRecordSet01 = null;
////				return;
////			ComboBox_Setting_Error:
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				oCombo = null;
////				//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				oRecordSet01 = null;
////				MDC_Com.MDC_GF_Message(ref "ComboBox_Setting_Error:" + Err().Number + " - " + Err().Description, ref "E");
////			}

////			public void CF_ChooseFromList()
////			{
////				// ERROR: Not supported in C#: OnErrorStatement

////				////ChooseFromList 설정
////				return;
////			CF_ChooseFromList_Error:
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				MDC_Com.MDC_GF_Message(ref "CF_ChooseFromList_Error:" + Err().Number + " - " + Err().Description, ref "E");
////			}

////			public void FormItemEnabled()
////			{
////				// ERROR: Not supported in C#: OnErrorStatement

////				if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
////				{
////				}
////				else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
////				{
////				}
////				else if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
////				{
////				}
////				return;
////			FormItemEnabled_Error:
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				MDC_Com.MDC_GF_Message(ref "FormItemEnabled_Error:" + Err().Number + " - " + Err().Description, ref "E");
////			}

////			public void Add_MatrixRow(int oRow, ref bool RowIserted = false)
////			{
////				// ERROR: Not supported in C#: OnErrorStatement

////				////행추가여부
////				if (RowIserted == false)
////				{
////					oDS_PS_CO120L.InsertRecord((oRow));
////				}
////				oMat01.AddRow();
////				oDS_PS_CO120L.Offset = oRow;
////				oDS_PS_CO120L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
////				oMat01.LoadFromDataSource();
////				return;
////			Add_MatrixRow_Error:
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				MDC_Com.MDC_GF_Message(ref "Add_MatrixRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
////			}

////			private void FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
////			{
////				// ERROR: Not supported in C#: OnErrorStatement

////				int i = 0;
////				short ErrNum = 0;
////				string sQry = null;
////				SAPbobsCOM.Recordset oRecordSet01 = null;
////				int sRow = 0;
////				string sSeq = null;

////				oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

////				sRow = oRow;

////				switch (oUID)
////				{
////					case "Mat01":
////						break;
////						//            If oCol = "CycleCod" Then
////						//                oForm01.Freeze True
////						//                oMat01.FlushToDataSource
////						//
////						//                If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("CycleCod").Cells(oRow).Specific.Value) <> "" Then
////						//                    oMat01.FlushToDataSource
////						//                    Call Add_MatrixRow(oMat01.RowCount, False)
////						//                    oMat01.Columns("CycleCod").Cells(oRow).Click ct_Regular
////						//                End If
////						//
////						//'                sQry = "Select ItemName, FrgnName From OITM Where ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(oRow).Specific.Value) & "'"
////						//'                oRecordSet01.DoQuery sQry
////						//'                oMat01.Columns("ItemName").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields(0).Value)
////						//'                oMat01.Columns("FrgnName").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields(1).Value)
////						//'
////						//'                oMat01.Columns("ItemCode").Cells(oRow).Click ct_Regular
////						//                oForm01.Freeze False
////						//            End If
////				}

////				//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				oRecordSet01 = null;
////				return;
////			FlushToItemValue_Error:
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				oRecordSet01 = null;
////				oForm01.Freeze(false);
////				if (ErrNum == 1)
////				{
////					MDC_Com.MDC_GF_Message(ref "구매견적문서가 취소되었거나 없습니다. 확인하세요.:" + Err().Number + " - " + Err().Description, ref "W");
////				}
////				else
////				{
////					MDC_Com.MDC_GF_Message(ref "FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
////				}
////			}

////			private bool HeaderSpaceLineDel()
////			{
////				bool functionReturnValue = false;
////				// ERROR: Not supported in C#: OnErrorStatement

////				short ErrNum = 0;
////				string DocNum = null;

////				ErrNum = 0;

////				//// Check
////				switch (true)
////				{
////					case string.IsNullOrEmpty(oDS_PS_CO120H.GetValue("U_YM", 0)):
////						ErrNum = 1;
////						goto HeaderSpaceLineDel_Error;
////						break;
////					case string.IsNullOrEmpty(oDS_PS_CO120H.GetValue("U_BPLId", 0)):
////						ErrNum = 2;
////						goto HeaderSpaceLineDel_Error;
////						break;
////				}

////				functionReturnValue = true;
////				return functionReturnValue;
////			HeaderSpaceLineDel_Error:
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				if (ErrNum == 1)
////				{
////					MDC_Com.MDC_GF_Message(ref "마감년월은 필수입력사항입니다. 확인하세요.", ref "E");
////				}
////				else if (ErrNum == 2)
////				{
////					MDC_Com.MDC_GF_Message(ref "사업장은 필수입력사항입니다. 확인하세요.", ref "E");
////				}
////				else
////				{
////					MDC_Com.MDC_GF_Message(ref "HeaderSpaceLineDel_Error:" + Err().Number + " - " + Err().Description, ref "E");
////				}
////				functionReturnValue = false;
////				return functionReturnValue;
////			}

////			private bool MatrixSpaceLineDel()
////			{
////				bool functionReturnValue = false;
////				// ERROR: Not supported in C#: OnErrorStatement

////				int i = 0;
////				short ErrNum = 0;
////				SAPbobsCOM.Recordset oRecordSet = null;
////				string sQry = null;

////				oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

////				ErrNum = 0;

////				oMat01.FlushToDataSource();

////				//// 라인
////				if (oMat01.VisualRowCount == 0)
////				{
////					ErrNum = 1;
////					goto MatrixSpaceLineDel_Error;
////				}
////				else if (oMat01.VisualRowCount == 1)
////				{
////					//        If oDS_PS_CO120L.GetValue("U_CycleCod", 0) = "" Then
////					//            ErrNum = 2
////					//            GoTo MatrixSpaceLineDel_Error
////					//        End If
////				}

////				for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
////				{
////					//        Select Case True
////					//            Case oDS_PS_CO120L.GetValue("U_ItemCode", i) = ""
////					//                ErrNum = 2
////					//                GoTo MatrixSpaceLineDel_Error
////					//            Case oDS_PS_CO120L.GetValue("U_Qty", i) = "" Or oDS_PS_CO120L.GetValue("U_Qty", i) = 0
////					//                ErrNum = 3
////					//                GoTo MatrixSpaceLineDel_Error
////					//            Case oDS_PS_CO120L.GetValue("U_Weight", i) = ""
////					//                ErrNum = 4
////					//                GoTo MatrixSpaceLineDel_Error
////					//            Case oDS_PS_CO120L.GetValue("U_Price", i) = 0
////					//                ErrNum = 5
////					//                GoTo MatrixSpaceLineDel_Error
////					//            Case oDS_PS_CO120L.GetValue("U_LinTotal", i) = 0
////					//                ErrNum = 6
////					//                GoTo MatrixSpaceLineDel_Error
////					//        End Select
////				}
////				oMat01.LoadFromDataSource();

////				//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				oRecordSet = null;
////				functionReturnValue = true;
////				return functionReturnValue;
////			MatrixSpaceLineDel_Error:
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				oRecordSet = null;
////				if (ErrNum == 1)
////				{
////					MDC_Com.MDC_GF_Message(ref "라인 데이터가 없습니다. 확인하세요.", ref "E");
////				}
////				else if (ErrNum == 2)
////				{
////					MDC_Com.MDC_GF_Message(ref "첫라인에 배부사이클 코드가 없습니다. 확인하세요.", ref "E");
////				}
////				else if (ErrNum == 3)
////				{
////					MDC_Com.MDC_GF_Message(ref "수량은 필수사항입니다. 확인하세요.", ref "E");
////				}
////				else if (ErrNum == 4)
////				{
////					MDC_Com.MDC_GF_Message(ref "중량은 필수사항입니다. 확인하세요.", ref "E");
////				}
////				else if (ErrNum == 5)
////				{
////					MDC_Com.MDC_GF_Message(ref "단가는 필수사항입니다. 확인하세요.", ref "E");
////				}
////				else if (ErrNum == 6)
////				{
////					MDC_Com.MDC_GF_Message(ref "금액은 필수사항입니다. 확인하세요.", ref "E");
////				}
////				else
////				{
////					MDC_Com.MDC_GF_Message(ref "MatrixSpaceLineDel_Error:" + Err().Number + " - " + Err().Description, ref "E");
////				}
////				functionReturnValue = false;
////				return functionReturnValue;
////			}

////			public void Delete_EmptyRow()
////			{
////				// ERROR: Not supported in C#: OnErrorStatement

////				int i = 0;

////				oMat01.FlushToDataSource();

////				for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
////				{
////					if (string.IsNullOrEmpty(Strings.Trim(oDS_PS_CO120L.GetValue("U_CycleCod", i))))
////					{
////						oDS_PS_CO120L.RemoveRecord(i);
////						//// Mat01에 마지막라인(빈라인) 삭제
////					}
////				}

////				oMat01.LoadFromDataSource();
////				return;
////			Delete_EmptyRow_Error:
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				MDC_Com.MDC_GF_Message(ref "Delete_EmptyRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
////			}

////			public void LoadData()
////			{
////				// ERROR: Not supported in C#: OnErrorStatement

////				short i = 0;
////				string sQry = null;
////				SAPbobsCOM.Recordset oRecordSet01 = null;
////				oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

////				string YM = null;
////				string BPLID = null;

////				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
////				YM = Strings.Trim(oForm01.Items.Item("YM").Specific.VALUE);
////				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
////				BPLID = Strings.Trim(oForm01.Items.Item("BPLId").Specific.VALUE);

////				oForm01.Freeze(true);
////				SAPbouiCOM.ProgressBar ProgBar01 = null;
////				ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

////				sQry = "EXEC [PS_CO120_01] '" + YM + "','" + BPLID + "'";
////				oRecordSet01.DoQuery(sQry);

////				oMat01.Clear();
////				oDS_PS_CO120L.Clear();

////				if ((oRecordSet01.RecordCount == 0))
////				{
////					MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.:" + Err().Number + " - " + Err().Description, ref "W");
////					//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////					oRecordSet01 = null;
////					oForm01.Freeze(false);
////					return;
////					//    ElseIf (oRecordSet01.Fields("ErrMessage").VALUE <> "") Then
////					//        Call Sbo_Application.MessageBox(oRecordSet01.Fields("ErrMessage").VALUE)
////					//        Set oRecordSet01 = Nothing
////					//        Call oForm01.Freeze(False)
////					//        Exit Sub
////				}

////				for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
////				{
////					if (i + 1 > oDS_PS_CO120L.Size)
////					{
////						oDS_PS_CO120L.InsertRecord((i));
////					}

////					oMat01.AddRow();
////					oDS_PS_CO120L.Offset = i;
////					oDS_PS_CO120L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
////					oDS_PS_CO120L.SetValue("U_POEntry", i, Strings.Trim(oRecordSet01.Fields.Item("POEntry").Value));
////					oDS_PS_CO120L.SetValue("U_POLine", i, Strings.Trim(oRecordSet01.Fields.Item("POLine").Value));
////					oDS_PS_CO120L.SetValue("U_Sequence", i, Strings.Trim(oRecordSet01.Fields.Item("Sequence").Value));
////					oDS_PS_CO120L.SetValue("U_ItemCode", i, Strings.Trim(oRecordSet01.Fields.Item("ItemCode").Value));
////					oDS_PS_CO120L.SetValue("U_ItemName", i, Strings.Trim(oRecordSet01.Fields.Item("ItemName").Value));
////					oDS_PS_CO120L.SetValue("U_CpCode", i, Strings.Trim(oRecordSet01.Fields.Item("CpCode").Value));
////					oDS_PS_CO120L.SetValue("U_CpName", i, Strings.Trim(oRecordSet01.Fields.Item("CpName").Value));
////					oDS_PS_CO120L.SetValue("U_CCCode", i, Strings.Trim(oRecordSet01.Fields.Item("CCCode").Value));
////					oDS_PS_CO120L.SetValue("U_CCName", i, Strings.Trim(oRecordSet01.Fields.Item("CCName").Value));
////					oDS_PS_CO120L.SetValue("U_ProdQty", i, Strings.Trim(oRecordSet01.Fields.Item("ProdQty").Value));
////					oDS_PS_CO120L.SetValue("U_DefQty", i, Strings.Trim(oRecordSet01.Fields.Item("DefQty").Value));
////					oDS_PS_CO120L.SetValue("U_Cost", i, Strings.Trim(oRecordSet01.Fields.Item("Cost").Value));
////					oDS_PS_CO120L.SetValue("U_Scrap", i, Strings.Trim(oRecordSet01.Fields.Item("Scrap").Value));
////					oDS_PS_CO120L.SetValue("U_Loss", i, Strings.Trim(oRecordSet01.Fields.Item("Loss").Value));

////					oRecordSet01.MoveNext();
////					ProgBar01.Value = ProgBar01.Value + 1;
////					ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
////				}
////				oMat01.LoadFromDataSource();
////				oMat01.AutoResizeColumns();
////				ProgBar01.Stop();
////				oForm01.Freeze(false);

////				//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				ProgBar01 = null;
////				//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				oRecordSet01 = null;
////				return;
////			LoadData_Error:
////				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////				ProgBar01.Stop();
////				oForm01.Freeze(false);
////				//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				ProgBar01 = null;
////				//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				oRecordSet01 = null;
////				MDC_Com.MDC_GF_Message(ref "LoadData_Error:" + Err().Number + " - " + Err().Description, ref "E");
////			}

////			public void SaveData()
////			{
////				//******************************************************************************
////				//Function ID : PS_CO120
////				//해당모듈    : SaveData
////				//기능        : 제품별원가계산 결과 저장
////				//인수        : 없음
////				//반환값      : 없음
////				//특이사항    : 없음
////				//******************************************************************************
////				// ERROR: Not supported in C#: OnErrorStatement


////				oForm01.Freeze(true);

////				int i = 0;
////				string sQry = null;

////				SAPbobsCOM.Recordset oRecordSet01 = null;
////				oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

////				string YM = null;
////				string BPLID = null;
////				string UserSign = null;

////				SAPbouiCOM.ProgressBar ProgBar01 = null;
////				ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("저장 중...", 100, false);

////				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
////				YM = Strings.Trim(oForm01.Items.Item("YM").Specific.VALUE);
////				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
////				BPLID = Strings.Trim(oForm01.Items.Item("BPLId").Specific.VALUE);
////				UserSign = Convert.ToString(SubMain.Sbo_Company.UserSignature);

////				sQry = "      EXEC [PS_CO120_50] '";
////				sQry = sQry + YM + "','";
////				sQry = sQry + BPLID + "','";
////				sQry = sQry + UserSign + "'";

////				oRecordSet01.DoQuery(sQry);

////				ProgBar01.Value = 100;
////				ProgBar01.Stop();
////				//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				ProgBar01 = null;

////				//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				oRecordSet01 = null;

////				oForm01.Freeze(false);
////				return;
////			SaveData_Error:


////				oForm01.Freeze(false);

////				ProgBar01.Value = 100;
////				ProgBar01.Stop();
////				//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				ProgBar01 = null;

////				//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
////				oRecordSet01 = null;

////				MDC_Com.MDC_GF_Message(ref "SaveData_Error:" + Err().Number + " - " + Err().Description, ref "E");

////			}

////			public void FindForm()
////			{
////				// ERROR: Not supported in C#: OnErrorStatement


////				string BPLID = null;
////				string StdYM = null;

////				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
////				BPLID = Strings.Trim(oForm01.Items.Item("BPLId").Specific.VALUE);
////				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
////				StdYM = oForm01.Items.Item("YM").Specific.VALUE;

////				//찾기모드 변경
////				oForm01.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

////				//UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
////				oForm01.Items.Item("BPLId").Specific.Select(BPLID);
////				//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
////				oForm01.Items.Item("YM").Specific.VALUE = StdYM;

////				oForm01.Items.Item("1").Click();

////				return;
////			FindFomr_Error:

////				MDC_Com.MDC_GF_Message(ref "FindForm_Error:" + Err().Number + " - " + Err().Description, ref "E");
////			}
////		}
////	}
