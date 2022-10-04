using System;
using System.IO;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;
using MsOutlook = Microsoft.Office.Interop.Outlook;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 급상여E-Mail 발송
    /// </summary>
    internal class PH_PY118 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY118A; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PH_PY118B; //등록라인

        private string oLastItemUID; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY118.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY118_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY118");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PH_PY118_CreateItems();
                PH_PY118_SetDocEntry();
                PH_PY118_FormItemEnabled();
            }
            catch (System.Exception ex)
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
        private void PH_PY118_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oDS_PH_PY118A = oForm.DataSources.DBDataSources.Item("@PH_PY118A");
                oDS_PH_PY118B = oForm.DataSources.DBDataSources.Item("@PH_PY118B");

                // 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                //사업장
                sQry = "SELECT BPLId, BPLName From[OBPL] order by 1";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("CLTCOD").Specific, "N");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //년월
                oForm.Items.Item("YM").Specific.Value = DateTime.Now.ToString("yyyyMM");

                //지급종류
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("1", "급여");
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("2", "상여");
                oForm.Items.Item("JOBTYP").DisplayDesc = true;
                oForm.Items.Item("JOBTYP").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                //지급구분
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P212' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JOBGBN").Specific, "N");
                oForm.Items.Item("JOBGBN").DisplayDesc = true;
                oForm.Items.Item("JOBGBN").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                //제목
                oForm.DataSources.UserDataSources.Add("Subject", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);

                //본문
                oForm.DataSources.UserDataSources.Add("Remark", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PH_PY118_SetDocEntry
        /// </summary>
        private void PH_PY118_SetDocEntry()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY118'", "");
                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.Value = "1";
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// PH_PY118_Add_MatrixRow
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        /// </summary>
        private void PH_PY118_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PH_PY118B.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PH_PY118B.Offset = oRow;
                oDS_PH_PY118B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY118_Add_MatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY118_MTX01
        /// </summary>
        private void PH_PY118_MTX01()
        {
            int i;
            string sQry;
            string errMessage = string.Empty;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string Param05;
            double Total = 0;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                Param02 = oForm.Items.Item("JOBGBN").Specific.Value.Trim();
                Param03 = oForm.Items.Item("JOBTYP").Specific.Value.Trim();
                Param04 = oForm.Items.Item("YM").Specific.Value;
                Param05 = oForm.Items.Item("JIGBIL").Specific.Value;

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                sQry = "EXEC PH_PY118_01 '" + Param01 + "','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + "'";
                oRecordSet.DoQuery(sQry);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();
                oDS_PH_PY118B.Clear(); //추가

                if (oRecordSet.RecordCount == 0)
                {
                    errMessage = "결과값이 존재하지않습니다.";
                    oMat01.Clear();
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY118B.Size)
                    {
                        oDS_PH_PY118B.InsertRecord(i);
                    }
                    oMat01.AddRow();

                    oDS_PH_PY118B.Offset = i;
                    oDS_PH_PY118B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY118B.SetValue("U_YM", i, oRecordSet.Fields.Item("YYMM").Value);
                    oDS_PH_PY118B.SetValue("U_MSTCOD", i, oRecordSet.Fields.Item("MSTCOD").Value);
                    oDS_PH_PY118B.SetValue("U_MSTNAM", i, oRecordSet.Fields.Item("MSTNAM").Value);
                    oDS_PH_PY118B.SetValue("U_SILJIG", i, oRecordSet.Fields.Item("SILJIG").Value);
                    oDS_PH_PY118B.SetValue("U_eMail", i, oRecordSet.Fields.Item("EmailAdress").Value);

                    if(!string.IsNullOrEmpty(oRecordSet.Fields.Item("EmailAdress").Value.ToString().Trim())) 
                    {
                        oDS_PH_PY118B.SetValue("U_Check", i, Convert.ToString('Y'));
                    }

                    Total += oRecordSet.Fields.Item("SILJIG").Value;
                    oRecordSet.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }
                string TotalSIL = String.Format("{0:#,###}", Total); //자릿값변환
                oForm.Items.Item("Total").Specific.Value = TotalSIL;
                TotalSIL = oForm.Items.Item("Total").Specific.Value;

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
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY118_MTX01:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY118_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    PH_PY118_SetDocEntry();
                    oForm.Items.Item("Btn02").Enabled = false;
                    oForm.Items.Item("Btn03").Enabled = false;
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                    oMat01.Columns.Item("Check").Editable = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("JOBGBN").Enabled = true;
                    oForm.Items.Item("JOBTYP").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("JIGBIL").Enabled = true;
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                    oMat01.Columns.Item("Check").Editable = true;
                    if (oForm.Items.Item("ControlYN").Specific.Value == "")
                    {
                        oForm.Items.Item("Btn02").Enabled = true;
                        oForm.Items.Item("Btn03").Enabled = false;
                    }
                    else if (oForm.Items.Item("ControlYN").Specific.Value == "S")
                    {
                        oForm.Items.Item("Btn02").Enabled = false;
                        oForm.Items.Item("Btn03").Enabled = true;
                    }
                    else if (oForm.Items.Item("ControlYN").Specific.Value == "C")
                    {
                        oForm.Items.Item("Btn02").Enabled = false;
                        oForm.Items.Item("Btn03").Enabled = false;
                    }
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("YM").Enabled = false;
                    oForm.Items.Item("JOBGBN").Enabled = false;
                    oForm.Items.Item("JOBTYP").Enabled = false;
                    oForm.Items.Item("JIGBIL").Enabled = false;
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                    oMat01.Columns.Item("Check").Editable = true;
                    if (oForm.Items.Item("ControlYN").Specific.Value == "")
                    {
                        oForm.Items.Item("Btn02").Enabled = true;
                        oForm.Items.Item("Btn03").Enabled = false;
                    }
                    else if (oForm.Items.Item("ControlYN").Specific.Value == "S")
                    {
                        oForm.Items.Item("Btn02").Enabled = false;
                        oForm.Items.Item("Btn03").Enabled = true;
                    }
                    else if (oForm.Items.Item("ControlYN").Specific.Value == "C")
                    {
                        oForm.Items.Item("Btn02").Enabled = false;
                        oForm.Items.Item("Btn03").Enabled = false;
                    }
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY118_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY118A_DataValidCheck()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                //년도
                if (string.IsNullOrEmpty(oForm.Items.Item("YM").Specific.Value.Trim()))
                {
                    errMessage = "년월은 필수입니다.";
                    throw new System.Exception();
                }

                //지급일자
                if (string.IsNullOrEmpty(oForm.Items.Item("JIGBIL").Specific.Value.Trim()))
                {
                    errMessage = "지급일자는 필수입니다.";
                    throw new System.Exception();
                }
                returnValue = true;
            }
            catch (System.Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                    return returnValue;
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY118_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
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

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                if (pVal.ItemUID == "Mat01")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID = pVal.ItemUID;
                        oLastColUID = pVal.ColUID;
                        oLastColRow = pVal.Row;
                    }
                }
                else
                {
                    oLastItemUID = pVal.ItemUID;
                    oLastColUID = "";
                    oLastColRow = 0;
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
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oMat01.SelectRow(pVal.Row, true, false);
                            }
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (System.Exception ex)
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
                     PH_PY118_FormItemEnabled();
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
        /// <summary>
        /// Raise_EVENT_KEY_DOWN
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.CharPressed == 9)
                {
                    if (pVal.ItemUID == "JIGBIL")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("JIGBIL").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_KEY_DOWN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY118A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY118B);
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
            int j = 0;
            string sQry;
            string sVersion;
            string sMSTCOD;
            string errMessage = string.Empty;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string Param05;
            SAPbouiCOM.ProgressBar ProgressBar01 = null; 
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                Param02 = oForm.Items.Item("JOBGBN").Specific.Value.Trim();
                Param03 = oForm.Items.Item("JOBTYP").Specific.Value.Trim();
                Param04 = oForm.Items.Item("YM").Specific.Value;
                Param05 = oForm.Items.Item("JIGBIL").Specific.Value;
                if (pVal.BeforeAction == true)
                {
                    sVersion = oForm.Items.Item("DocEntry").Specific.Value;

                    //조회
                    if (pVal.ItemUID == "Btn01")
                    {

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY118A_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PH_PY118_MTX01();
                        }
                    }

                    //추가
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY118A_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }

                            if (oDS_PH_PY118B.Size < 2)
                            {
                                errMessage = "조회 누르르고 추가하세오!";
                                BubbleEvent = false;
                                throw new System.Exception();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        {
                        }
                    }

                    //PDF생성
                    if (pVal.ItemUID == "Btn02")
                    {
                        ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("PDF 파일 생성 시작!", 50, false);
                        oMat01.FlushToDataSource();
                       
                        sQry = "DELETE FROM Z_TEMP_GONGTABLE"; 
                        oRecordSet.DoQuery(sQry);

                        sQry = "DELETE FROM Z_TEMP_JIGTABLE";
                        oRecordSet.DoQuery(sQry);

                        sQry = "EXEC PH_PY118_02 '" + Param01 + "','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + "'";
                        oRecordSet.DoQuery(sQry);

                        sQry = "EXEC PH_PY118_03 '" + Param01 + "','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + "'";
                        oRecordSet.DoQuery(sQry);

                        if (Param02 == "1" && Param03 == "1")
                        {
                            sQry = "EXEC PH_PY118_04 '" + Param01 + "','" + Param04 + "'";
                            oRecordSet.DoQuery(sQry);
                            j = 1;
                        }

                        for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            if (oDS_PH_PY118B.GetValue("U_Check", i).ToString().Trim() == "Y")
                            {
                                if (oDS_PH_PY118B.GetValue("U_SaveYN", i).ToString().Trim() != "Y")
                                {
                                    if (!string.IsNullOrEmpty(oDS_PH_PY118B.GetValue("U_eMail", i).ToString().Trim()))
                                    {
                                        sMSTCOD = oDS_PH_PY118B.GetValue("U_MSTCOD", i).ToString().Trim();
                                        if (Make_PDF_File(sMSTCOD, sVersion) == false)
                                        {
                                            errMessage = "PDF저장이 완료되지 않았습니다.";
                                            throw new Exception();
                                        }
                                    }
                                }
                            }
                            ProgressBar01.Value += 1;
                            ProgressBar01.Text = ProgressBar01.Value + "/" + (oMat01.VisualRowCount) + "건 PDF 파일 생성 중...!";
                        }

                        ProgressBar01.Stop();
                        
                        if(j==1)
                        {
                            sQry = "DELETE FROM Z_PH_PY118_011";
                            oRecordSet.DoQuery(sQry);
                            j = 0;
                        }

                        sQry = "Update [@PH_PY118A] Set U_ControlYN = 'S' where DocEntry = '" + sVersion + "'";
                        oRecordSet.DoQuery(sQry);

                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        PH_PY118_FormItemEnabled();
                        oForm.Items.Item("DocEntry").Specific.Value = sVersion;
                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }

                    //메일보내기
                    if (pVal.ItemUID == "Btn03")
                    {
                        ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("eMail 메일전송", 50, false);
                        oMat01.FlushToDataSource();
                        for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            if (oDS_PH_PY118B.GetValue("U_Check", i).ToString().Trim() == "Y")
                            {
                                if (oDS_PH_PY118B.GetValue("U_SendYN", i).ToString().Trim() != "Y")
                                {
                                    if (!string.IsNullOrEmpty(oDS_PH_PY118B.GetValue("U_SaveYN", i).ToString().Trim()))
                                    {
                                        sMSTCOD = oDS_PH_PY118B.GetValue("U_MSTCOD", i).ToString().Trim();
                                        if (Send_EMail(sMSTCOD, sVersion) == false)//사번
                                        {
                                            errMessage = "전송 중 오류가 발생했습니다.";
                                            throw new Exception();
                                        }
                                    }
                                }
                            }
                            ProgressBar01.Value += 1;
                            ProgressBar01.Text = ProgressBar01.Value + "/" + (oMat01.VisualRowCount) + "건 eMail전송중...!";
                        }
                        ProgressBar01.Stop();

                        sQry = "Update [@PH_PY118A] Set U_ControlYN = 'C' Where DocEntry = '" + sVersion + "'";
                        oRecordSet.DoQuery(sQry);

                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        PH_PY118_FormItemEnabled();
                        oForm.Items.Item("DocEntry").Specific.Value = sVersion;
                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }

                    else if (pVal.BeforeAction == false)
                    {
                        if (pVal.ItemUID == "1")
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY118_FormItemEnabled();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// <summary>
        /// Make_PDF_File
        /// </summary>
        /// <param name="p_MSTCOD">사번</param>
        /// <param name="p_Version">문서번호</param>
        /// <returns></returns>
        private bool Make_PDF_File(string p_MSTCOD, string p_Version)
        {
            bool ReturnValue = false;
            string WinTitle;
            string ReportName;
            string sQry;
            string STDYER;
            string STDMON;
            string Main_Folder;
            string Sub_Folder1;
            string Sub_Folder2;
            string Sub_Folder3;
            string CLTCOD;
            string JOBTYP;
            string JOBGBN;
            string YM;
            string JIGBIL;
            string ExportString;
            string psgovID;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                JOBGBN = oForm.Items.Item("JOBGBN").Specific.Value.Trim();
                JOBTYP = oForm.Items.Item("JOBTYP").Specific.Value.Trim();
                YM = oForm.Items.Item("YM").Specific.Value;
                JIGBIL = oForm.Items.Item("JIGBIL").Specific.Value;
                STDYER = YM.Substring(0, 4);
                STDMON = YM.Substring(4, 2);

                sQry = "Select RIGHT(U_govID,7) From [@PH_PY001A] WHERE Code = '" + p_MSTCOD + "'";
                oRecordSet01.DoQuery(sQry);
                psgovID = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                WinTitle = "개인별급여명세서_" + p_MSTCOD;
                ReportName = "PH_PY118_01.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //레포트 그대로날리는변수 
                List<PSH_DataPackClass> dataPackSub1ReportParameter = new List<PSH_DataPackClass>(); //서브레포트 그대로날리는변수 
                List<PSH_DataPackClass> dataPackSub2ReportParameter = new List<PSH_DataPackClass>(); //서브레포트 그대로날리는변수 

                //본문레포트
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@JOBGBN", JOBGBN)); //지급구분
                dataPackParameter.Add(new PSH_DataPackClass("@JOBTYP", JOBTYP)); //지급종류
                dataPackParameter.Add(new PSH_DataPackClass("@YM", YM)); //년월
                dataPackParameter.Add(new PSH_DataPackClass("@JIGBIL", JIGBIL)); //지급일
                dataPackParameter.Add(new PSH_DataPackClass("@MSTCOD", p_MSTCOD)); //사번

                //서브레포트1
                dataPackSub1ReportParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD, "SUB_PH_PY118_01"));
                dataPackSub1ReportParameter.Add(new PSH_DataPackClass("@JOBGBN", JOBGBN, "SUB_PH_PY118_01"));
                dataPackSub1ReportParameter.Add(new PSH_DataPackClass("@JOBTYP", JOBTYP, "SUB_PH_PY118_01"));
                dataPackSub1ReportParameter.Add(new PSH_DataPackClass("@YM", YM, "SUB_PH_PY118_01"));
                dataPackSub1ReportParameter.Add(new PSH_DataPackClass("@JIGBIL", JIGBIL, "SUB_PH_PY118_01"));
                dataPackSub1ReportParameter.Add(new PSH_DataPackClass("@MSTCOD", p_MSTCOD, "SUB_PH_PY118_01"));

                //서브레포트2
                dataPackSub2ReportParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD, "SUB_PH_PY118_02"));
                dataPackSub2ReportParameter.Add(new PSH_DataPackClass("@JOBGBN", JOBGBN, "SUB_PH_PY118_02"));
                dataPackSub2ReportParameter.Add(new PSH_DataPackClass("@JOBTYP", JOBTYP, "SUB_PH_PY118_02"));
                dataPackSub2ReportParameter.Add(new PSH_DataPackClass("@YM", YM, "SUB_PH_PY118_02"));
                dataPackSub2ReportParameter.Add(new PSH_DataPackClass("@JIGBIL", JIGBIL, "SUB_PH_PY118_02"));
                dataPackSub2ReportParameter.Add(new PSH_DataPackClass("@MSTCOD", p_MSTCOD, "SUB_PH_PY118_02"));

                Main_Folder = @"C:\PSH_개인별급여명세서";
                Sub_Folder1 = @"C:\PSH_개인별급여명세서\" + STDYER + "";
                Sub_Folder2 = @"C:\PSH_개인별급여명세서\" + STDYER + @"\" + STDMON + "";
                Sub_Folder3 = @"C:\PSH_개인별급여명세서\" + STDYER + @"\" + STDMON + @"\" + "문서번호" + p_Version + "";

                //디렉토리 생성
                Dir_Exists(Main_Folder);
                Dir_Exists(Sub_Folder1);
                Dir_Exists(Sub_Folder2);
                Dir_Exists(Sub_Folder3);

                ExportString = Sub_Folder3 + @"\" + p_MSTCOD + "_개인별급여명세서_" + STDYER + "" + STDMON + ".pdf";

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackSub1ReportParameter, dataPackSub2ReportParameter, ExportString);
                
                // Open an existing document. Providing an unrequired password is ignored.
                PdfDocument document = PdfReader.Open(ExportString, PdfDocumentOpenMode.Modify);

                PdfSecuritySettings securitySettings = document.SecuritySettings;

                securitySettings.UserPassword = "manager";   //개개인암호
                securitySettings.OwnerPassword = psgovID;    //마스터암호

                // Restrict some rights.
                securitySettings.PermitAccessibilityExtractContent = false;
                securitySettings.PermitAnnotations = false;
                securitySettings.PermitAssembleDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitFormsFill = true;
                securitySettings.PermitFullQualityPrint = false;
                securitySettings.PermitModifyDocument = true;
                securitySettings.PermitPrint = false;

                // PDF문서 저장
                document.Save(ExportString);

                sQry = "Update [@PH_PY118B] Set U_SaveYN = 'Y' Where U_MSTCOD = '" + p_MSTCOD + "' And DocEntry = '" + p_Version + "'";
                oRecordSet01.DoQuery(sQry);

                ReturnValue = true;
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Make_PDF_File_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
            return ReturnValue;
        }

        /// <summary>
        /// 디렉토리 체크, 폴더 생성
        /// </summary>
        /// <param name="strDirName">경로</param>
        /// <returns></returns>
        private int Dir_Exists(string strDirName)
        {
            int ReturnValue = 0;

            try
            {
                DirectoryInfo di = new DirectoryInfo(strDirName); //DirectoryInfo 생성
                //DirectoryInfo.Exists로 폴더 존재유무 확인
                if (di.Exists)
                {
                    ReturnValue = 1;
                }
                else
                {
                    di.Create();
                    ReturnValue = 0;
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Make_PDF_File_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
            return ReturnValue;
        }

        /// <summary>
        /// Send_EMail
        /// </summary>
        /// <param name="p_MSTCOD"></param>
        /// <param name="p_Version"></param>
        /// <returns></returns>
        private bool Send_EMail(string p_MSTCOD, string p_Version)
        {
            bool ReturnValue = false;
            string strToAddress;
            string strSubject;
            string strBody;
            string Sub_Folder3;
            string sQry;
            string YM;
            string STDYER;
            string STDMON;
            string MSTCOD;
            string Version;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                MSTCOD = p_MSTCOD;
                Version = p_Version;
                YM = oForm.Items.Item("YM").Specific.Value;
                STDYER = YM.Substring(0, 4);
                STDMON = YM.Substring(4, 2);

                Sub_Folder3 = @"C:\PSH_개인별급여명세서\" + STDYER + @"\" + STDMON + @"\" + "문서번호" + Version + "";

                sQry = "Select U_Subject, U_Remark From [@PH_PY118A] Where Docentry = '" + Version + "'";
                oRecordSet01.DoQuery(sQry);
                strSubject = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                strBody = oRecordSet01.Fields.Item(1).Value.ToString().Trim();

                sQry = "Select U_eMail From [@PH_PY118B] Where U_MSTCOD = '" + MSTCOD + "' AND Docentry = '" + Version + "'";
                oRecordSet01.DoQuery(sQry);
                strToAddress = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

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
                MsOutlook.Attachment oAttach = mail.Attachments.Add(Sub_Folder3 + @"\" + p_MSTCOD + "_개인별급여명세서_" + STDYER + "" + STDMON + ".pdf");
                mail.Send();

                mail = null;
                outlookApp = null;

                sQry = "Update [@PH_PY118B] Set U_SendYN = 'Y' Where U_MSTCOD = '" + MSTCOD + "' And DocEntry = '" + Version + "'";
                oRecordSet01.DoQuery(sQry);

                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(Sub_Folder3 + @"\" + p_MSTCOD + "_개인별급여명세서_" + STDYER + "" + STDMON + ".pdf");

                //원래코드시작
                //SmtpClient smtp = new SmtpClient("smtp.naver.com");
                //SmtpClient smtp = new SmtpClient("pscsn.poongsan.co.kr");
                //SmtpClient smtp = new SmtpClient("smtp.office365.com");
                //SmtpClient smtp = new SmtpClient("smtp.gmail.com");

                //smtp.Port = 587; //네이버
                //smtp.Port = 25; //풍산
                //smtp.UseDefaultCredentials = true;
                //smtp.EnableSsl = true;
                //smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                //smtp.Timeout = 20000;

                //smtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;  //Naver 인 경우
                //smtp.Credentials = new NetworkCredential("2220501", "p2220501!"); //address, PW
                //smtp.Credentials = new NetworkCredential("wgpark@poongsan.co.kr", "1q2w3e4r)*"); //address, PW
                //smtp.Credentials = new NetworkCredential("dakkorea1@gmail.com", "dak440310*"); //address, PW

                //smtp.Send(mail);
                //원래코드 끝

                ReturnValue = true;
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Send_EMail_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            return ReturnValue;
        }

        /// FormMenuEvent
        /// <summary>
        /// 메뉴이벤트
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
                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
                                }

                                oMat01.FlushToDataSource();
                                oDS_PH_PY118B.RemoveRecord(oDS_PH_PY118B.Size - 1);
                                oMat01.Clear();
                                oMat01.LoadFromDataSource();

                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("LineNum").Cells.Item(oMat01.RowCount).Specific.Value))
                                {
                                    PH_PY118_AddMatrixRow(oMat01.RowCount, false);
                                }
                            }
                            break;
                        case "1281": //찾기
                            PH_PY118_FormItemEnabled();
                            break;
                        case "1282": //추가
                            PH_PY118_FormItemEnabled();
                            PH_PY118_SetDocEntry();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PH_PY118_FormItemEnabled();
                            break;
                        case "1287": //복제
                            break;
                    }
                }
            }
            catch (System.Exception ex)
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

