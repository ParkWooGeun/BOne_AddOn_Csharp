//using Microsoft.VisualBasic;
//using Microsoft.VisualBasic.Compatibility;
//using System;
//using System.Collections;
//using System.Data;
//using System.Diagnostics;
//using System.Drawing;
//using System.Windows.Forms;
// // ERROR: Not supported in C#: OptionDeclaration
using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 급상여은행파일생성
    /// </summary>
    internal class PH_PY101 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY101A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY101B;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;
        
        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY101.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY101_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY101");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                PH_PY101_CreateItems();
                PH_PY101_EnableMenus();
                PH_PY101_SetDocument(oFromDocEntry01);
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PH_PY101_CreateItems()
        {
            string sQry = string.Empty;
            int i = 0;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oDS_PH_PY101A = oForm.DataSources.DBDataSources.Item("@PH_PY101A");
                oDS_PH_PY101B = oForm.DataSources.DBDataSources.Item("@PH_PY101B");
                
                oMat1 = oForm.Items.Item("Mat1").Specific;
                ////@PH_PY101B
                
                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                oForm.AutoManaged = true;
                //MDC_SetMod.AutoManaged(ref oForm, ref "Code");
                ////----------------------------------------------------------------------------------------------
                //// 기본사항
                ////----------------------------------------------------------------------------------------------

                //// 구분
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P180' AND U_UseYN= 'Y'";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oMat1.Columns.Item("SIType").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                        oRecordSet.MoveNext();
                    }
                }
                oMat1.Columns.Item("SIType").DisplayDesc = true;
                
                //// 기준금액
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P181' AND U_UseYN= 'Y'";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oMat1.Columns.Item("BaseType").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                        oRecordSet.MoveNext();
                    }
                }
                oMat1.Columns.Item("BaseType").DisplayDesc = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY101_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY101_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY101_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY101_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PH_PY101_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY101_FormItemEnabled();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY119_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        public void PH_PY101_AddMatrixRow()
        {
            int oRow = 0;
            try
            {
                oForm.Freeze(true);

                ////[Mat1]
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY101B.GetValue("U_SIType", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY101B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY101B.InsertRecord((oRow));
                        }
                        oDS_PH_PY101B.Offset = oRow;
                        oDS_PH_PY101B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY101B.SetValue("U_SIType", oRow, "");
                        oDS_PH_PY101B.SetValue("U_EChgRate", oRow, Convert.ToString(0));
                        oDS_PH_PY101B.SetValue("U_CChgRate", oRow, Convert.ToString(0));
                        oDS_PH_PY101B.SetValue("U_TChgRate", oRow, Convert.ToString(0));
                        oDS_PH_PY101B.SetValue("U_BaseType", oRow, "");
                        oDS_PH_PY101B.SetValue("U_LowLimit", oRow, Convert.ToString(0));
                        oDS_PH_PY101B.SetValue("U_UppLimit", oRow, Convert.ToString(0));
                        oDS_PH_PY101B.SetValue("U_LineMemo", oRow, "");
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY101B.Offset = oRow - 1;
                        oDS_PH_PY101B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY101B.SetValue("U_SIType", oRow - 1, "");
                        oDS_PH_PY101B.SetValue("U_EChgRate", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY101B.SetValue("U_CChgRate", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY101B.SetValue("U_TChgRate", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY101B.SetValue("U_BaseType", oRow - 1, "");
                        oDS_PH_PY101B.SetValue("U_LowLimit", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY101B.SetValue("U_UppLimit", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY101B.SetValue("U_LineMemo", oRow - 1, "");
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY101B.Offset = oRow;
                    oDS_PH_PY101B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY101B.SetValue("U_SIType", oRow, "");
                    oDS_PH_PY101B.SetValue("U_EChgRate", oRow, Convert.ToString(0));
                    oDS_PH_PY101B.SetValue("U_CChgRate", oRow, Convert.ToString(0));
                    oDS_PH_PY101B.SetValue("U_TChgRate", oRow, Convert.ToString(0));
                    oDS_PH_PY101B.SetValue("U_BaseType", oRow, "");
                    oDS_PH_PY101B.SetValue("U_LowLimit", oRow, Convert.ToString(0));
                    oDS_PH_PY101B.SetValue("U_UppLimit", oRow, Convert.ToString(0));
                    oDS_PH_PY101B.SetValue("U_LineMemo", oRow, "");
                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY101_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY101_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {

                    oForm.EnableMenu("1281", true);
                    ////문서찾기
                    oForm.EnableMenu("1282", false);
                    ////문서추가

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {

                    oForm.EnableMenu("1281", false);
                    ////문서찾기
                    oForm.EnableMenu("1282", true);
                    ////문서추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {

                    oForm.EnableMenu("1281", true);
                    ////문서찾기
                    oForm.EnableMenu("1282", true);
                    ////문서추가

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY101_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
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
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY101_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }

                            ////해야할일 작업
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY101_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                            ////해야할일 작업

                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
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
                                PH_PY101_FormItemEnabled();
                                PH_PY101_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY101_FormItemEnabled();
                                PH_PY101_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY101_FormItemEnabled();
                            }
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
                oForm.Freeze(false);
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
            string sQry = string.Empty;
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
                        PH_PY101_AddMatrixRow();
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    oMat1.LoadFromDataSource();
                    PH_PY101_FormItemEnabled();
                    PH_PY101_AddMatrixRow();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY101A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY101B);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY101_DataValidCheck()
        {
            bool functionReturnValue = false;

            functionReturnValue = false;
            int i = 0;
            string sQry = string.Empty;
            try
            {
                ////헤더
                if (string.IsNullOrEmpty(oDS_PH_PY101A.GetValue("Code", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("년월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                //// 라인 ---------------------------
                if (oMat1.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                    {
                        //구분
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("SIType").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("구분은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("SIType").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            functionReturnValue = false;
                            return functionReturnValue;
                        }

                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                oMat1.FlushToDataSource();
                //// Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY101B.Size > 1)
                    oDS_PH_PY101B.RemoveRecord((oDS_PH_PY101B.Size - 1));

                oMat1.LoadFromDataSource();

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY119_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
            return functionReturnValue;
        }

        /// <summary>
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            int i = 0;
            try
            {
                if ((pVal.BeforeAction == true))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            if (PSH_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1284":
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
                            break;
                    }
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY101_FormItemEnabled();
                            PH_PY101_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281":
                            ////문서찾기
                            PH_PY101_FormItemEnabled();
                            PH_PY101_AddMatrixRow();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282":
                            ////문서추가
                            PH_PY101_FormItemEnabled();
                            PH_PY101_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY101_FormItemEnabled();
                            break;
                        case "1293":
                            //// 행삭제
                            //// [MAT1] 급여 수당
                            if (oMat1.RowCount != oMat1.VisualRowCount)
                            {
                                oMat1.FlushToDataSource();

                                while ((i <= oDS_PH_PY101B.Size - 1))
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY101B.GetValue("U_FILD01", i)))
                                    {
                                        oDS_PH_PY101B.RemoveRecord((i));
                                        i = 0;
                                    }
                                    else
                                    {
                                        i = i + 1;
                                    }
                                }

                                for (i = 0; i <= oDS_PH_PY101B.Size; i++)
                                {
                                    oDS_PH_PY101B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                }

                                oMat1.LoadFromDataSource();
                            }
                            PH_PY101_AddMatrixRow();
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
    }
}

//namespace MDC_HR_Addon
//{
//    internal class PH_PY101
//    {
//        ////********************************************************************************
//        ////  File           : PH_PY101.cls
//        ////  Module         : 급여관리 > 급여관리
//        ////  Desc           : 보험율등록
//        ////********************************************************************************

//        public string oFormUniqueID;
//        public SAPbouiCOM.Form oForm;

//        public SAPbouiCOM.Matrix oMat1;

//        private SAPbouiCOM.DBDataSource oDS_PH_PY101A;
//        private SAPbouiCOM.DBDataSource oDS_PH_PY101B;

//        private string oLastItemUID;
//        private string oLastColUID;
//        private int oLastColRow;

//        public void LoadForm(string oFromDocEntry01 = "")
//        {

//            int i = 0;
//            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//            // ERROR: Not supported in C#: OnErrorStatement


//            oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY101.srf");
//            oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//            for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
//            {
//                oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//                oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//            }
//            oFormUniqueID = "PH_PY101_" + GetTotalFormsCount();
//            SubMain.AddForms(this, oFormUniqueID, "PH_PY101");
//            MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//            oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//            oForm.SupportedModes = -1;
//            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//            oForm.DataBrowser.BrowseBy = "Code";

//            oForm.Freeze(true);
//            PH_PY101_CreateItems();
//            PH_PY101_EnableMenus();
//            PH_PY101_SetDocument(oFromDocEntry01);
//            //    Call PH_PY101_FormResize

//            oForm.Update();
//            oForm.Freeze(false);

//            oForm.Visible = true;
//            //UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oXmlDoc = null;
//            return;
//        LoadForm_Error:

//            oForm.Update();
//            oForm.Freeze(false);
//            //UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oXmlDoc = null;
//            //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oForm = null;
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private bool PH_PY101_CreateItems()
//        {
//            bool functionReturnValue = false;

//            string sQry = null;
//            int i = 0;

//            SAPbouiCOM.EditText oEdit = null;
//            SAPbouiCOM.ComboBox oCombo = null;
//            SAPbouiCOM.Column oColumn = null;
//            SAPbouiCOM.Columns oColumns = null;

//            SAPbobsCOM.Recordset oRecordSet = null;

//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.Freeze(true);

//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            oDS_PH_PY101A = oForm.DataSources.DBDataSources("@PH_PY101A");
//            oDS_PH_PY101B = oForm.DataSources.DBDataSources("@PH_PY101B");


//            oMat1 = oForm.Items.Item("Mat1").Specific;
//            ////@PH_PY101B


//            oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//            oMat1.AutoResizeColumns();

//            oForm.AutoManaged = true;
//            MDC_SetMod.AutoManaged(ref oForm, ref "Code");
//            ////----------------------------------------------------------------------------------------------
//            //// 기본사항
//            ////----------------------------------------------------------------------------------------------

//            //// 구분
//            oColumn = oMat1.Columns.Item("SIType");
//            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P180' AND U_UseYN= 'Y'";
//            oRecordSet.DoQuery(sQry);
//            if (oRecordSet.RecordCount > 0)
//            {
//                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
//                {
//                    oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
//                    oRecordSet.MoveNext();
//                }
//            }
//            oColumn.DisplayDesc = true;


//            //// 기준금액
//            oColumn = oMat1.Columns.Item("BaseType");
//            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P181' AND U_UseYN= 'Y'";
//            oRecordSet.DoQuery(sQry);
//            if (oRecordSet.RecordCount > 0)
//            {
//                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
//                {
//                    oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
//                    oRecordSet.MoveNext();
//                }
//            }
//            oColumn.DisplayDesc = true;

//            //UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oEdit = null;
//            //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCombo = null;
//            //UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oColumn = null;
//            //UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oColumns = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            oForm.Freeze(false);
//            return functionReturnValue;
//        PH_PY101_CreateItems_Error:

//            //UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oEdit = null;
//            //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCombo = null;
//            //UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oColumn = null;
//            //UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oColumns = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            oForm.Freeze(false);
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY101_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }

//        private void PH_PY101_EnableMenus()
//        {

//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.EnableMenu("1283", true);
//            ////제거
//            oForm.EnableMenu("1284", false);
//            ////취소
//            oForm.EnableMenu("1293", true);
//            ////행삭제

//            return;
//        PH_PY101_EnableMenus_Error:

//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY101_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void PH_PY101_SetDocument(string oFromDocEntry01)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            if ((string.IsNullOrEmpty(oFromDocEntry01)))
//            {
//                PH_PY101_FormItemEnabled();
//                PH_PY101_AddMatrixRow();
//            }
//            else
//            {
//                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//                PH_PY101_FormItemEnabled();
//                //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("Code").Specific.Value = oFromDocEntry01;
//                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//            }
//            return;
//        PH_PY101_SetDocument_Error:

//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY101_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void PH_PY101_FormItemEnabled()
//        {
//            SAPbouiCOM.ComboBox oCombo = null;

//            // ERROR: Not supported in C#: OnErrorStatement



//            oForm.Freeze(true);
//            if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
//            {

//                oForm.EnableMenu("1281", true);
//                ////문서찾기
//                oForm.EnableMenu("1282", false);
//                ////문서추가

//            }
//            else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
//            {

//                oForm.EnableMenu("1281", false);
//                ////문서찾기
//                oForm.EnableMenu("1282", true);
//                ////문서추가
//            }
//            else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
//            {

//                oForm.EnableMenu("1281", true);
//                ////문서찾기
//                oForm.EnableMenu("1282", true);
//                ////문서추가

//            }
//            oForm.Freeze(false);
//            return;
//        PH_PY101_FormItemEnabled_Error:

//            oForm.Freeze(false);
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY101_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }


//        public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            string sQry = null;
//            int i = 0;
//            SAPbouiCOM.ComboBox oCombo = null;
//            SAPbobsCOM.Recordset oRecordSet = null;

//            // ERROR: Not supported in C#: OnErrorStatement


//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            switch (pVal.EventType)
//            {
//                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//                    ////1

//                    if (pVal.BeforeAction == true)
//                    {
//                        if (pVal.ItemUID == "1")
//                        {
//                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                            {
//                                if (PH_PY101_DataValidCheck() == false)
//                                {
//                                    BubbleEvent = false;
//                                }

//                                ////해야할일 작업
//                            }
//                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                            {
//                                if (PH_PY101_DataValidCheck() == false)
//                                {
//                                    BubbleEvent = false;
//                                }
//                                ////해야할일 작업

//                            }
//                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//                            {
//                            }
//                        }
//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        if (pVal.ItemUID == "1")
//                        {
//                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                            {
//                                if (pVal.ActionSuccess == true)
//                                {
//                                    PH_PY101_FormItemEnabled();
//                                    PH_PY101_AddMatrixRow();
//                                }
//                            }
//                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                            {
//                                if (pVal.ActionSuccess == true)
//                                {
//                                    PH_PY101_FormItemEnabled();
//                                    PH_PY101_AddMatrixRow();
//                                }
//                            }
//                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//                            {
//                                if (pVal.ActionSuccess == true)
//                                {
//                                    PH_PY101_FormItemEnabled();
//                                }
//                            }
//                        }
//                    }
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//                    ////2
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//                    ////3
//                    switch (pVal.ItemUID)
//                    {
//                        case "Mat1":
//                            if (pVal.Row > 0)
//                            {
//                                oLastItemUID = pVal.ItemUID;
//                                oLastColUID = pVal.ColUID;
//                                oLastColRow = pVal.Row;
//                            }
//                            break;
//                        default:
//                            oLastItemUID = pVal.ItemUID;
//                            oLastColUID = "";
//                            oLastColRow = 0;
//                            break;
//                    }
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//                    ////4
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//                    ////5
//                    oForm.Freeze(true);
//                    if (pVal.BeforeAction == true)
//                    {

//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        if (pVal.ItemChanged == true)
//                        {
//                            PH_PY101_AddMatrixRow();
//                        }
//                    }
//                    oForm.Freeze(false);
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_CLICK:
//                    ////6
//                    if (pVal.BeforeAction == true)
//                    {
//                        switch (pVal.ItemUID)
//                        {
//                            case "Mat1":
//                                if (pVal.Row > 0)
//                                {
//                                    oMat1.SelectRow(pVal.Row, true, false);
//                                }
//                                break;
//                        }

//                        switch (pVal.ItemUID)
//                        {
//                            case "Mat1":
//                                if (pVal.Row > 0)
//                                {
//                                    oLastItemUID = pVal.ItemUID;
//                                    oLastColUID = pVal.ColUID;
//                                    oLastColRow = pVal.Row;
//                                }
//                                break;
//                            default:
//                                oLastItemUID = pVal.ItemUID;
//                                oLastColUID = "";
//                                oLastColRow = 0;
//                                break;
//                        }
//                    }
//                    else if (pVal.BeforeAction == false)
//                    {

//                    }
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//                    ////7
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//                    ////8
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
//                    ////9
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//                    ////10
//                    oForm.Freeze(true);
//                    if (pVal.BeforeAction == true)
//                    {

//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        if (pVal.ItemChanged == true)
//                        {

//                        }
//                    }
//                    oForm.Freeze(false);
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//                    ////11
//                    if (pVal.BeforeAction == true)
//                    {
//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        oMat1.LoadFromDataSource();

//                        PH_PY101_FormItemEnabled();
//                        PH_PY101_AddMatrixRow();

//                    }
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
//                    ////12
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
//                    ////16
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//                    ////17
//                    if (pVal.BeforeAction == true)
//                    {
//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        SubMain.RemoveForms(oFormUniqueID);
//                        //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oForm = null;
//                        //UPGRADE_NOTE: oDS_PH_PY101A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oDS_PH_PY101A = null;
//                        //UPGRADE_NOTE: oDS_PH_PY101B 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oDS_PH_PY101B = null;

//                        //UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oMat1 = null;

//                    }
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//                    ////18
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//                    ////19
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
//                    ////20
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//                    ////21
//                    if (pVal.BeforeAction == true)
//                    {

//                    }
//                    else if (pVal.BeforeAction == false)
//                    {

//                    }
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
//                    ////22
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
//                    ////23
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//                    ////27
//                    if (pVal.BeforeAction == true)
//                    {

//                    }
//                    else if (pVal.Before_Action == false)
//                    {
//                        //                If pVal.ItemUID = "Code" Then
//                        //                    Call MDC_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY101A", "Code")
//                        //                End If
//                    }
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
//                    ////37
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
//                    ////38
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_Drag:
//                    ////39
//                    break;

//            }

//            //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCombo = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;

//            return;
//        Raise_FormItemEvent_Error:
//            ///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//            oForm.Freeze((false));
//            //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCombo = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }


//        public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
//        {
//            int i = 0;
//            // ERROR: Not supported in C#: OnErrorStatement

//            oForm.Freeze(true);

//            if ((pVal.BeforeAction == true))
//            {
//                switch (pVal.MenuUID)
//                {
//                    case "1283":
//                        if (MDC_Globals.Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
//                        {
//                            BubbleEvent = false;
//                            return;
//                        }
//                        break;
//                    case "1284":
//                        break;
//                    case "1286":
//                        break;
//                    case "1293":
//                        break;
//                    case "1281":
//                        break;
//                    case "1282":
//                        break;
//                    case "1288":
//                    case "1289":
//                    case "1290":
//                    case "1291":
//                        break;
//                }
//            }
//            else if ((pVal.BeforeAction == false))
//            {
//                switch (pVal.MenuUID)
//                {
//                    case "1283":
//                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                        PH_PY101_FormItemEnabled();
//                        PH_PY101_AddMatrixRow();
//                        break;
//                    case "1284":
//                        break;
//                    case "1286":
//                        break;
//                    //            Case "1293":
//                    //                Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
//                    case "1281":
//                        ////문서찾기
//                        PH_PY101_FormItemEnabled();
//                        PH_PY101_AddMatrixRow();
//                        oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                        break;
//                    case "1282":
//                        ////문서추가
//                        PH_PY101_FormItemEnabled();
//                        PH_PY101_AddMatrixRow();
//                        break;
//                    case "1288":
//                    case "1289":
//                    case "1290":
//                    case "1291":
//                        PH_PY101_FormItemEnabled();
//                        break;
//                    case "1293":
//                        //// 행삭제
//                        //// [MAT1] 급여 수당
//                        if (oMat1.RowCount != oMat1.VisualRowCount)
//                        {
//                            oMat1.FlushToDataSource();

//                            while ((i <= oDS_PH_PY101B.Size - 1))
//                            {
//                                if (string.IsNullOrEmpty(oDS_PH_PY101B.GetValue("U_FILD01", i)))
//                                {
//                                    oDS_PH_PY101B.RemoveRecord((i));
//                                    i = 0;
//                                }
//                                else
//                                {
//                                    i = i + 1;
//                                }
//                            }

//                            for (i = 0; i <= oDS_PH_PY101B.Size; i++)
//                            {
//                                oDS_PH_PY101B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//                            }

//                            oMat1.LoadFromDataSource();
//                        }
//                        PH_PY101_AddMatrixRow();
//                        break;
//                }
//            }
//            oForm.Freeze(false);
//            return;
//        Raise_FormMenuEvent_Error:
//            oForm.Freeze(false);
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//        {

//            // ERROR: Not supported in C#: OnErrorStatement


//            if ((BusinessObjectInfo.BeforeAction == true))
//            {
//                switch (BusinessObjectInfo.EventType)
//                {
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//                        ////33
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//                        ////34
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//                        ////35
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//                        ////36
//                        break;
//                }
//            }
//            else if ((BusinessObjectInfo.BeforeAction == false))
//            {
//                switch (BusinessObjectInfo.EventType)
//                {
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//                        ////33
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//                        ////34
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//                        ////35
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//                        ////36
//                        break;
//                }
//            }
//            return;
//        Raise_FormDataEvent_Error:


//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//        }

//        public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
//        {

//            // ERROR: Not supported in C#: OnErrorStatement


//            if (pVal.BeforeAction == true)
//            {
//            }
//            else if (pVal.BeforeAction == false)
//            {
//            }
//            switch (pVal.ItemUID)
//            {
//                case "Mat1":
//                    if (pVal.Row > 0)
//                    {
//                        oLastItemUID = pVal.ItemUID;
//                        oLastColUID = pVal.ColUID;
//                        oLastColRow = pVal.Row;
//                    }
//                    break;
//                default:
//                    oLastItemUID = pVal.ItemUID;
//                    oLastColUID = "";
//                    oLastColRow = 0;
//                    break;
//            }
//            return;
//        Raise_RightClickEvent_Error:

//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void PH_PY101_AddMatrixRow()
//        {
//            int oRow = 0;

//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.Freeze(true);

//            ////[Mat1]
//            oMat1.FlushToDataSource();
//            oRow = oMat1.VisualRowCount;

//            if (oMat1.VisualRowCount > 0)
//            {
//                if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY101B.GetValue("U_SIType", oRow - 1))))
//                {
//                    if (oDS_PH_PY101B.Size <= oMat1.VisualRowCount)
//                    {
//                        oDS_PH_PY101B.InsertRecord((oRow));
//                    }
//                    oDS_PH_PY101B.Offset = oRow;
//                    oDS_PH_PY101B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//                    oDS_PH_PY101B.SetValue("U_SIType", oRow, "");
//                    oDS_PH_PY101B.SetValue("U_EChgRate", oRow, Convert.ToString(0));
//                    oDS_PH_PY101B.SetValue("U_CChgRate", oRow, Convert.ToString(0));
//                    oDS_PH_PY101B.SetValue("U_TChgRate", oRow, Convert.ToString(0));
//                    oDS_PH_PY101B.SetValue("U_BaseType", oRow, "");
//                    oDS_PH_PY101B.SetValue("U_LowLimit", oRow, Convert.ToString(0));
//                    oDS_PH_PY101B.SetValue("U_UppLimit", oRow, Convert.ToString(0));
//                    oDS_PH_PY101B.SetValue("U_LineMemo", oRow, "");
//                    oMat1.LoadFromDataSource();
//                }
//                else
//                {
//                    oDS_PH_PY101B.Offset = oRow - 1;
//                    oDS_PH_PY101B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//                    oDS_PH_PY101B.SetValue("U_SIType", oRow - 1, "");
//                    oDS_PH_PY101B.SetValue("U_EChgRate", oRow - 1, Convert.ToString(0));
//                    oDS_PH_PY101B.SetValue("U_CChgRate", oRow - 1, Convert.ToString(0));
//                    oDS_PH_PY101B.SetValue("U_TChgRate", oRow - 1, Convert.ToString(0));
//                    oDS_PH_PY101B.SetValue("U_BaseType", oRow - 1, "");
//                    oDS_PH_PY101B.SetValue("U_LowLimit", oRow - 1, Convert.ToString(0));
//                    oDS_PH_PY101B.SetValue("U_UppLimit", oRow - 1, Convert.ToString(0));
//                    oDS_PH_PY101B.SetValue("U_LineMemo", oRow - 1, "");
//                    oMat1.LoadFromDataSource();
//                }
//            }
//            else if (oMat1.VisualRowCount == 0)
//            {
//                oDS_PH_PY101B.Offset = oRow;
//                oDS_PH_PY101B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//                oDS_PH_PY101B.SetValue("U_SIType", oRow, "");
//                oDS_PH_PY101B.SetValue("U_EChgRate", oRow, Convert.ToString(0));
//                oDS_PH_PY101B.SetValue("U_CChgRate", oRow, Convert.ToString(0));
//                oDS_PH_PY101B.SetValue("U_TChgRate", oRow, Convert.ToString(0));
//                oDS_PH_PY101B.SetValue("U_BaseType", oRow, "");
//                oDS_PH_PY101B.SetValue("U_LowLimit", oRow, Convert.ToString(0));
//                oDS_PH_PY101B.SetValue("U_UppLimit", oRow, Convert.ToString(0));
//                oDS_PH_PY101B.SetValue("U_LineMemo", oRow, "");
//                oMat1.LoadFromDataSource();
//            }

//            oForm.Freeze(false);
//            return;
//        PH_PY101_AddMatrixRow_Error:
//            oForm.Freeze(false);
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY101_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void PH_PY101_FormClear()
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            string DocEntry = null;
//            //UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY101'", ref "");
//            if (Convert.ToDouble(DocEntry) == 0)
//            {
//                //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("DocEntry").Specific.Value = 1;
//            }
//            else
//            {
//                //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
//            }
//            return;
//        PH_PY101_FormClear_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY101_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public bool PH_PY101_DataValidCheck()
//        {
//            bool functionReturnValue = false;
//            // ERROR: Not supported in C#: OnErrorStatement

//            functionReturnValue = false;
//            int i = 0;
//            string sQry = null;
//            SAPbobsCOM.Recordset oRecordSet = null;

//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            ////헤더
//            if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY101A.GetValue("Code", 0))))
//            {
//                MDC_Globals.Sbo_Application.SetStatusBarMessage("사원번호는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                functionReturnValue = false;
//                return functionReturnValue;
//            }

//            //// 라인 ---------------------------
//            if (oMat1.VisualRowCount > 1)
//            {
//                for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
//                {
//                    ////구분
//                    //UPGRADE_WARNING: oMat1.Columns(SIType).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    if (string.IsNullOrEmpty(oMat1.Columns.Item("SIType").Cells.Item(i).Specific.Value))
//                    {
//                        MDC_Globals.Sbo_Application.SetStatusBarMessage("구분은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                        oMat1.Columns.Item("SIType").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                        functionReturnValue = false;
//                        return functionReturnValue;
//                    }

//                }
//            }
//            else
//            {
//                MDC_Globals.Sbo_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                functionReturnValue = false;
//                return functionReturnValue;
//            }

//            oMat1.FlushToDataSource();
//            //// Matrix 마지막 행 삭제(DB 저장시)
//            if (oDS_PH_PY101B.Size > 1)
//                oDS_PH_PY101B.RemoveRecord((oDS_PH_PY101B.Size - 1));

//            oMat1.LoadFromDataSource();

//            functionReturnValue = true;
//            return functionReturnValue;


//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//        PH_PY101_DataValidCheck_Error:


//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            functionReturnValue = false;
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY101_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }

//        private void PH_PY101_MTX01()
//        {

//            ////메트릭스에 데이터 로드

//            int i = 0;
//            string sQry = null;

//            string Param01 = null;
//            string Param02 = null;
//            string Param03 = null;
//            string Param04 = null;

//            SAPbobsCOM.Recordset oRecordSet = null;

//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.Freeze(true);
//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Param01 = oForm.Items.Item("Param01").Specific.Value;
//            //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Param02 = oForm.Items.Item("Param01").Specific.Value;
//            //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Param03 = oForm.Items.Item("Param01").Specific.Value;
//            //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Param04 = oForm.Items.Item("Param01").Specific.Value;

//            sQry = "SELECT 10";
//            oRecordSet.DoQuery(sQry);

//            oMat1.Clear();
//            oMat1.FlushToDataSource();
//            oMat1.LoadFromDataSource();

//            if ((oRecordSet.RecordCount == 0))
//            {
//                MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
//                goto PH_PY101_MTX01_Exit;
//            }

//            SAPbouiCOM.ProgressBar ProgressBar01 = null;
//            ProgressBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

//            for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
//            {
//                if (i != 0)
//                {
//                    oDS_PH_PY101B.InsertRecord((i));
//                }
//                oDS_PH_PY101B.Offset = i;
//                oDS_PH_PY101B.SetValue("U_COL01", i, oRecordSet.Fields.Item(0).Value);
//                oDS_PH_PY101B.SetValue("U_COL02", i, oRecordSet.Fields.Item(1).Value);
//                oRecordSet.MoveNext();
//                ProgressBar01.Value = ProgressBar01.Value + 1;
//                ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
//            }
//            oMat1.LoadFromDataSource();
//            oMat1.AutoResizeColumns();
//            oForm.Update();

//            ProgressBar01.Stop();
//            //UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            ProgressBar01 = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            oForm.Freeze(false);
//            return;
//        PH_PY101_MTX01_Exit:
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            oForm.Freeze(false);
//            if ((ProgressBar01 != null))
//            {
//                ProgressBar01.Stop();
//            }
//            return;
//        PH_PY101_MTX01_Error:
//            ProgressBar01.Stop();
//            //UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            ProgressBar01 = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            oForm.Freeze(false);
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY101_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public bool PH_PY101_Validate(string ValidateType)
//        {
//            bool functionReturnValue = false;
//            // ERROR: Not supported in C#: OnErrorStatement

//            functionReturnValue = true;
//            object i = null;
//            int j = 0;
//            string sQry = null;
//            SAPbobsCOM.Recordset oRecordSet = null;
//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            //UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY101A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY101A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
//            {
//                MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                functionReturnValue = false;
//                goto PH_PY101_Validate_Exit;
//            }
//            //
//            if (ValidateType == "수정")
//            {

//            }
//            else if (ValidateType == "행삭제")
//            {

//            }
//            else if (ValidateType == "취소")
//            {

//            }
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            return functionReturnValue;
//        PH_PY101_Validate_Exit:
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            return functionReturnValue;
//        PH_PY101_Validate_Error:
//            functionReturnValue = false;
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY101_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }
//    }
//}
