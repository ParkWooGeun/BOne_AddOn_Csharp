using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 정산징수및환급대장(집계)
    /// </summary>
    internal class PH_PYA55 : PSH_BaseClass
    {
        public string oFormUniqueID01;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFromDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PYA55.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PYA55_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PYA55");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PYA55_CreateItems();
                PH_PYA55_FormItemEnabled();
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
                oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PH_PYA55_CreateItems()
        {
            try
            {
                oForm.Freeze(true);

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 년도
                oForm.DataSources.UserDataSources.Add("YYYY", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("YYYY").Specific.DataBind.SetBound(true, "", "YYYY");
                oForm.DataSources.UserDataSources.Item("YYYY").Value = Convert.ToString(DateTime.Now.Year - 1);

                // 재직구분
                oForm.DataSources.UserDataSources.Add("Div", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Div").Specific.ValidValues.Add("1", "전체");
                oForm.Items.Item("Div").Specific.ValidValues.Add("2", "재직자");
                oForm.Items.Item("Div").Specific.ValidValues.Add("3", "퇴직자");
                oForm.Items.Item("Div").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PYA55_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        public void PH_PYA55_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
                oForm.Items.Item("CLTCOD").Specific.ValidValues.Add("%", "전사업장");
                oForm.Items.Item("CLTCOD").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PYA55_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Raise_FormItemEvent
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">이벤트 </param>
        /// <param name="BubbleEvent">Bubble Event</param>
        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                    ////2
                    break;


                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                    ////3

                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                    ////4
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                                                              // Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                    ////7
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                    ////8
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
                    ////9
                    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                                                             //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
                    ////12
                    break;


                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                    ////16
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
                    ////18
                    break;


                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
                    ////19
                    break;


                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                    ////20
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                                                             // Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
                    ////22
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
                    ////23
                    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                                                                  // Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
                    ////37
                    break;

                case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
                    ////38
                    break;

                case SAPbouiCOM.BoEventTypes.et_Drag:
                    ////39
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn01")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PYA55_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
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
                    SubMain.Remove_Forms(oFormUniqueID01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
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
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PH_PYA55_Print_Report01()
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;

            string CLTCOD = string.Empty;
            string YYYY = string.Empty;
            string Div = string.Empty;

            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.Value.ToString().Trim();
            YYYY = oForm.Items.Item("YYYY").Specific.Value.ToString().Trim();
            Div = oForm.Items.Item("Div").Specific.Selected.Value.ToString().Trim();

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                WinTitle = "[PH_PYA55] 정산징수및환급대장(집계)";
                ReportName = "PH_PYA55_01.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter List
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List

                //Formula
                if (CLTCOD == "%")
                {
                    dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", "전사업장")); //전사업장
                }
                else
                {
                    dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //사업장
                }

                dataPackFormula.Add(new PSH_DataPackClass("@YYYY", YYYY));

                if (Div == "1")
                {
                    dataPackFormula.Add(new PSH_DataPackClass("@Div", "전체"));
                }
                else if (Div == "2")
                {
                    dataPackFormula.Add(new PSH_DataPackClass("@Div", "재직자"));
                }
                else if (Div == "3")
                {
                    dataPackFormula.Add(new PSH_DataPackClass("@Div", "퇴직자"));
                }

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@saup", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@yyyy", YYYY));
                dataPackParameter.Add(new PSH_DataPackClass("@Div", Div));

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PYA55_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}



//using Microsoft.VisualBasic;
//using Microsoft.VisualBasic.Compatibility;
//using System;
//using System.Collections;
//using System.Data;
//using System.Diagnostics;
//using System.Drawing;
//using System.Windows.Forms;
////using Microsoft.Office.Interop;
//using SAPbobsCOM;
//using SAPbouiCOM;
//using PSH_BOne_AddOn.Data;

//namespace PSH_BOne_AddOn
//{
//    internal class PH_PYA55 : PSH_BaseClass
//    {
//        ////********************************************************************************
//        ////  File           : PH_PYA55.cls
//        ////  Module         : PH
//        ////  Desc           : 정산징수및환급대장(집계)
//        ////  작성자         : NGY
//        ////  DATE           : 2016.01.19
//        ////********************************************************************************

//        public string oFormUniqueID;
//        //public SAPbouiCOM.Form oForm;

//        //'// 그리드 사용시
//        //Public oGrid1           As SAPbouiCOM.Grid
//        //Public oDS_PH_PYA55     As SAPbouiCOM.DataTable
//        //
//        //'// 매트릭스 사용시
//        //Public oMat1 As SAPbouiCOM.Matrix
//        //Private oDS_PH_PYA55A As SAPbouiCOM.DBDataSource
//        //Private oDS_PH_PYA55B As SAPbouiCOM.DBDataSource

//        private string oLastItemUID;
//        private string oLastColUID;
//        private int oLastColRow;

//        public override void LoadForm(string oFormDocEntry01)
//        {

//            int i = 0;
//            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//            try
//            {
//                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PYA55.srf");
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue.ToString() + "_" + (SubMain.Get_TotalFormsCount());
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

//                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
//                {
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//                }

//                oFormUniqueID = "PH_PYA55_" + SubMain.Get_TotalFormsCount();
//                SubMain.Add_Forms(this, oFormUniqueID, "PH_PYA55");

//                string strXml = string.Empty;
//                strXml = oXmlDoc.xml.ToString();

//                PSH_Globals.SBO_Application.LoadBatchActions(ref strXml);
//                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

//                oForm.SupportedModes = -1;
//                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                //    oForm.DataBrowser.BrowseBy = "Code"

//                oForm.Freeze(true);
//                PH_PYA55_CreateItems();
//                PH_PYA55_EnableMenus();
//                PH_PYA55_SetDocument(oFormDocEntry01);
//                //    Call PH_PYA55_FormResize
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("Form_Load Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            }
//            finally
//            {
//                oForm.Update();
//                oForm.Freeze(false);
//                oForm.Visible = true;
//                //메모리 해제
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
//            }
//        }

//        private void PH_PYA55_CreateItems()
//        {
//            string sQry = string.Empty;
//            SAPbouiCOM.ComboBox oCombo = null;

//            SAPbobsCOM.Recordset oRecordSet = null;

//            oForm.Freeze(true);

//            try
//            {
//                PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();
//                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//                ////사업장
//                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//                oCombo = oForm.Items.Item("CLTCOD").Specific;
//                oCombo.DataBind.SetBound(true, "", "CLTCOD");

//                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//                oRecordSet.DoQuery(sQry);
//                oCombo.ValidValues.Add("%", "전 사업장");
//                while (!(oRecordSet.EoF))
//                {
//                    oCombo.ValidValues.Add(Strings.Trim(oRecordSet.Fields.Item(0).Value), Strings.Trim(oRecordSet.Fields.Item(1).Value));
//                    oRecordSet.MoveNext();
//                }
//                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);


//                //// 접속자에 따른 권한별 사업장 콤보박스세팅
//                //Call CLTCOD_Select(oForm, "CLTCOD")

//                oForm.Items.Item("YYYY").Specific.String = Convert.ToDouble(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY")) - 1;

//                ////재직구분
//                oCombo = oForm.Items.Item("Div").Specific;
//                oCombo.ValidValues.Add("1", "전체");
//                oCombo.ValidValues.Add("2", "재직자");
//                oCombo.ValidValues.Add("3", "퇴직자");
//                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//                ////커서를 첫번째 ITEM으로 지정
//                oForm.ActiveItem = "CLTCOD";
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PYA55_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCombo); //메모리 해제
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
//            }         
//        }

//        private void PH_PYA55_EnableMenus()
//        {
//            try
//            {
//                oForm.EnableMenu("1283", true);
//                ////제거
//                oForm.EnableMenu("1284", false);
//                ////취소
//                oForm.EnableMenu("1293", true);
//                ////행삭제
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PYA55_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        private void PH_PYA55_SetDocument(string oFromDocEntry01)
//        {
//            try
//            {
//                if ((string.IsNullOrEmpty(oFromDocEntry01)))
//                {
//                    PH_PYA55_FormItemEnabled();
//                    //PH_PYA55_AddMatrixRow();
//                }
//                else
//                {
//                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//                    PH_PYA55_FormItemEnabled();

//                    oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
//                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                }
//            }
//            catch(Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PYA55_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        public void PH_PYA55_FormItemEnabled()
//        {
//            try
//            {
//                oForm.Freeze(true);
//                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
//                {
//                    oForm.EnableMenu("1281", true);
//                    ////문서찾기
//                    oForm.EnableMenu("1282", false);
//                    ////문서추가
//                }
//                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
//                {
//                    oForm.EnableMenu("1281", false);
//                    ////문서찾기
//                    oForm.EnableMenu("1282", true);
//                    ////문서추가
//                }
//                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
//                {
//                    oForm.EnableMenu("1281", true);
//                    ////문서찾기
//                    oForm.EnableMenu("1282", true);
//                    ////문서추가
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PYA55_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//        {
//            try
//            {
//                switch (pval.EventType)
//                {
//                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//                    ////1

//                    if (pval.BeforeAction == true)
//                    {
//                        if (pval.ItemUID == "Btn01")
//                        {
//                            //PH_PYA55_DataValidCheck();
//                            PH_PYA55_Print_Report01();
//                        }
//                    }
//                    else if (pval.BeforeAction == false)
//                    {


//                    }
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//                    ////2
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//                    ////3
//                    switch (pval.ItemUID)
//                    {
//                        case "Mat01":
//                        case "Grid01":
//                            if (pval.Row > 0)
//                            {
//                                oLastItemUID = pval.ItemUID;
//                                oLastColUID = pval.ColUID;
//                                oLastColRow = pval.Row;
//                            }
//                            break;
//                        default:
//                            oLastItemUID = pval.ItemUID;
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
//                    if (pval.BeforeAction == true)
//                    {



//                    }
//                    else if (pval.BeforeAction == false)
//                    {
//                        if (pval.ItemChanged == true)
//                        {
//                            switch (pval.ItemUID)
//                            {


//                            }
//                        }
//                    }

//                    oForm.Freeze(false);
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_CLICK:
//                    ////6
//                    if (pval.BeforeAction == true)
//                    {
//                        switch (pval.ItemUID)
//                        {
//                            case "Mat01":
//                                break;
//                                //                    If pval.Row > 0 Then
//                                //                        Call oMat1.SelectRow(pval.Row, True, False)
//                                //                    End If
//                        }

//                        switch (pval.ItemUID)
//                        {
//                            case "Mat01":
//                            case "Grid01":
//                                if (pval.Row > 0)
//                                {
//                                    oLastItemUID = pval.ItemUID;
//                                    oLastColUID = pval.ColUID;
//                                    oLastColRow = pval.Row;
//                                }
//                                break;
//                            default:
//                                oLastItemUID = pval.ItemUID;
//                                oLastColUID = "";
//                                oLastColRow = 0;
//                                break;
//                        }
//                    }
//                    else if (pval.BeforeAction == false)
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
//                    if (pval.BeforeAction == true)
//                    {

//                    }
//                    else if (pval.BeforeAction == false)
//                    {
//                        if (pval.ItemChanged == true)
//                        {
//                            switch (pval.ItemUID)
//                            {


//                            }
//                        }
//                    }
//                    oForm.Freeze(false);
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//                    ////11
//                    if (pval.BeforeAction == true)
//                    {
//                    }
//                    else if (pval.BeforeAction == false)
//                    {
//                        //oMat1.LoadFromDataSource

//                        PH_PYA55_FormItemEnabled();
//                        //PH_PYA55_AddMatrixRow();

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
//                    if (pval.BeforeAction == true)
//                    {
//                    }
//                    else if (pval.BeforeAction == false)
//                    {
//                        SubMain.Remove_Forms(oFormUniqueID);
//                        //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oForm = null;
//                        //                Set oDS_PH_PYA55A = Nothing
//                        //                Set oDS_PH_PYA55B = Nothing

//                        //Set oMat1 = Nothing
//                        //Set oGrid1 = Nothing

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
//                    if (pval.BeforeAction == true)
//                    {

//                    }
//                    else if (pval.BeforeAction == false)
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
//                    if (pval.BeforeAction == true)
//                    {

//                    }
//                    else if (pval.Before_Action == false)
//                    {

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

//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PYA55_Raise_ItemEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }            
//        }

//        public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//        {
//            oForm.Freeze(true);
//            try
//            {
//                if ((pval.BeforeAction == true))
//                {
//                    switch (pval.MenuUID)
//                    {
//                        case "1283":
//                            if (PSH_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
//                            {
//                                BubbleEvent = false;
//                                return;
//                            }
//                            break;
//                        case "1284":
//                            break;
//                        case "1286":
//                            break;
//                        case "1293":
//                            break;
//                        case "1281":
//                            break;
//                        case "1282":
//                            break;
//                        case "1288":
//                        case "1289":
//                        case "1290":
//                        case "1291":
//                            break;
//                    }
//                }
//                else if ((pval.BeforeAction == false))
//                {
//                    switch (pval.MenuUID)
//                    {
//                        case "1283":
//                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                            PH_PYA55_FormItemEnabled();
//                            //PH_PYA55_AddMatrixRow();
//                            break;
//                        case "1284":
//                            break;
//                        case "1286":
//                            break;
//                        //            Case "1293":
//                        //                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//                        case "1281":
//                            ////문서찾기
//                            PH_PYA55_FormItemEnabled();
//                            //PH_PYA55_AddMatrixRow();
//                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                            break;
//                        case "1282":
//                            ////문서추가
//                            PH_PYA55_FormItemEnabled();
//                            //PH_PYA55_AddMatrixRow();
//                            break;
//                        case "1288":
//                        case "1289":
//                        case "1290":
//                        case "1291":
//                            PH_PYA55_FormItemEnabled();
//                            break;
//                        case "1293":
//                            //// 행삭제
//                            //                '// [MAT1 용]
//                            //                 If oMat1.RowCount <> oMat1.VisualRowCount Then
//                            //                    oMat1.FlushToDataSource
//                            //
//                            //                    While (i <= oDS_PH_PYA55B.Size - 1)
//                            //                        If oDS_PH_PYA55B.GetValue("U_FILD01", i) = "" Then
//                            //                            oDS_PH_PYA55B.RemoveRecord (i)
//                            //                            i = 0
//                            //                        Else
//                            //                            i = i + 1
//                            //                        End If
//                            //                    Wend
//                            //
//                            //                    For i = 0 To oDS_PH_PYA55B.Size
//                            //                        Call oDS_PH_PYA55B.setValue("U_LineNum", i, i + 1)
//                            //                    Next i
//                            //
//                            //                    oMat1.LoadFromDataSource
//                            //End If
//                            //PH_PYA55_AddMatrixRow();
//                            break;
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PYA55_Raise_MenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//        {
//            try
//            {
//                if ((BusinessObjectInfo.BeforeAction == true))
//                {
//                    switch (BusinessObjectInfo.EventType)
//                    {
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//                            ////33
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//                            ////34
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//                            ////35
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//                            ////36
//                            break;
//                    }
//                }
//                else if ((BusinessObjectInfo.BeforeAction == false))
//                {
//                    switch (BusinessObjectInfo.EventType)
//                    {
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//                            ////33
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//                            ////34
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//                            ////35
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//                            ////36
//                            break;
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PYA55_Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pval.BeforeAction == true)
//                {
//                }
//                else if (pval.BeforeAction == false)
//                {
//                }
//                switch (pval.ItemUID)
//                {
//                    case "Mat01":
//                        if (pval.Row > 0)
//                        {
//                            oLastItemUID = pval.ItemUID;
//                            oLastColUID = pval.ColUID;
//                            oLastColRow = pval.Row;
//                        }
//                        break;
//                    default:
//                        oLastItemUID = pval.ItemUID;
//                        oLastColUID = "";
//                        oLastColRow = 0;
//                        break;
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PYA55_Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }        
//        }

//        public void PH_PYA55_FormClear()
//        {
//            string DocEntry = string.Empty;

//            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                DocEntry = DataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PYA55'", "");
//                if (Convert.ToDouble(DocEntry) == 0)
//                {                    
//                    oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//                }
//                else
//                {                 
//                    oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PYA55_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        public bool PH_PYA55_Validate(string ValidateType)
//        {
//            bool functionReturnValue = false;
//            functionReturnValue = false;

//            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                if (DataHelpClass.GetValue("SELECT Canceled FROM [@PH_PYA55A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y")
//                {
//                    PSH_Globals.SBO_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                    functionReturnValue = false;
//                    goto PH_PYA55_Validate_Exit;
//                }
//                //
//                if (ValidateType == "수정")
//                {

//                }
//                else if (ValidateType == "행삭제")
//                {

//                }
//                else if (ValidateType == "취소")
//                {

//                }

//                functionReturnValue = true;
//            }
//            catch (Exception ex)
//            {
//                functionReturnValue = false;

//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PYA55_Validate_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }

//            return functionReturnValue;            
//        }

//        private void PH_PYA55_Print_Report01()
//        {

//            string DocNum = null;
//            short ErrNum = 0;
//            string WinTitle = null;
//            string ReportName = null;
//            string sQry = null;
//            SAPbobsCOM.Recordset oRecordSet = null;

//            // ERROR: Not supported in C#: OnErrorStatement


//            string CLTCOD = null;
//            string yyyy = null;
//            string Div = null;


//            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


//            /// ODBC 연결 체크
//            if (ConnectODBC() == false)
//            {
//                goto PH_PYA55_Print_Report01_Error;
//            }


//            ////인자 MOVE , Trim 시키기..
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            yyyy = Strings.Trim(oForm.Items.Item("YYYY").Specific.VALUE);
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Div = Strings.Trim(oForm.Items.Item("Div").Specific.VALUE);


//            /// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

//            WinTitle = "[PH_PYA55] 정산징수및환급대장(집계)";
//            ReportName = "PH_PYA55_01.rpt";
//            PSH_Globals.gRpt_Formula = new string[4];
//            PSH_Globals.gRpt_Formula_Value = new string[4];

//            /// Formula 수식필드

//            PSH_Globals.gRpt_Formula[1] = "CLTCOD";
//            sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y' AND U_Code = '" + CLTCOD + "'";
//            oRecordSet.DoQuery(sQry);
//            if (CLTCOD == "%")
//            {
//                PSH_Globals.gRpt_Formula_Value[1] = "전 사업장";
//            }
//            else
//            {
//                //UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                PSH_Globals.gRpt_Formula_Value[1] = oRecordSet.Fields.Item(0).Value;
//            }

//            PSH_Globals.gRpt_Formula[2] = "YYYY";
//            PSH_Globals.gRpt_Formula_Value[2] = yyyy;

//            PSH_Globals.gRpt_Formula[3] = "Div";
//            if (Div == "1")
//            {
//                PSH_Globals.gRpt_Formula_Value[3] = "전체";
//            }
//            else if (Div == "2")
//            {
//                PSH_Globals.gRpt_Formula_Value[3] = "재직자";
//            }
//            else if (Div == "3")
//            {
//                PSH_Globals.gRpt_Formula_Value[3] = "퇴직자";
//            }
//            PSH_Globals.gRpt_SRptSqry = new string[2];
//            PSH_Globals.gRpt_SRptName = new string[2];
//            PSH_Globals.gRpt_SFormula = new string[2, 2];
//            PSH_Globals.gRpt_SFormula_Value = new string[2, 2];
//            /// SubReport


//            /// Procedure 실행"
//            sQry = "EXEC [PH_PYA55_01] '" + CLTCOD + "', '" + yyyy + "', '" + Div + "'";

//            //    oRecordSet.DoQuery sQry
//            //    If oRecordSet.RecordCount = 0 Then
//            //        ErrNum = 1
//            //        GoTo PH_PYA55_Print_Report01_Error
//            //    End If

//            if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V", , 2) == false)
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            }

//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            return;
//        PH_PYA55_Print_Report01_Error:

//            if (ErrNum == 1)
//            {
//                //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                oRecordSet = null;
//                MDC_Com.MDC_GF_Message(ref "출력할 데이터가 없습니다. 확인해 주세요.", ref "E");
//            }
//            else
//            {
//                //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                oRecordSet = null;
//                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PYA55_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            }

//        }
//    }
//}