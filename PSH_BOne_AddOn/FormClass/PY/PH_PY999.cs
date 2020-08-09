using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 사용자 권한 등록
    /// </summary>
    internal class PH_PY999 : PSH_BaseClass
    {
        public string oFormUniqueID01;

        //'// 그리드 사용시
        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.DataTable oDS_PH_PY999A;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFromDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY999.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY999_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY999");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                oForm.Freeze(true);
                PH_PY999_CreateItems();
                PH_PY999_FormItemEnabled();
                PH_PY999_EnableMenus();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
         //       oForm.Visible = true;
         //       oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PH_PY999_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY999");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY999");
                oDS_PH_PY999A = oForm.DataSources.DataTables.Item("PH_PY999");
                
                // 구분
                oForm.Items.Item("pGubun").Specific.ValidValues.Add("B", "기본");
                oForm.Items.Item("pGubun").Specific.ValidValues.Add("H", "인사");
                oForm.Items.Item("pGubun").DisplayDesc = true;

                // 폴더/화면구분
                oForm.Items.Item("pFSGubun").Specific.ValidValues.Add("F", "폴더");
                oForm.Items.Item("pFSGubun").Specific.ValidValues.Add("S", "화면");
                oForm.Items.Item("pFSGubun").Specific.ValidValues.Add("C", "복제");
                oForm.Items.Item("pFSGubun").DisplayDesc = true;

                ////복제 YN
                //oForm.Items.Item("CopyYN").Specific.ValOff = "N";
                //oForm.Items.Item("CopyYN").Specific.ValOn = "Y";

                //// 급여변동자료적용
                //oForm.DataSources.UserDataSources.Add("CopyYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                //oForm.Items.Item("CopyYN").Specific.ValOn = "Y";
                //oForm.Items.Item("CopyYN").Specific.ValOff = "N";
                //oForm.Items.Item("CopyYN").Specific.DataBind.SetBound(true, "", "CopyYN");
                //oForm.DataSources.UserDataSources.Item("CopyYN").Value = "N";

                // 순서
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='A006'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Modual").Specific,  "Y");
                oForm.Items.Item("Modual").DisplayDesc = true;

                // Position
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='A005'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Position").Specific, "Y");
                oForm.Items.Item("Position").DisplayDesc = true;

                // Sub1
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Sub1").Specific, "Y");
                oForm.Items.Item("Sub1").DisplayDesc = true;

                // Sub2
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Sub2").Specific, "Y");
                oForm.Items.Item("Sub2").DisplayDesc = true;

                // Sub3
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Sub3").Specific, "Y");
                oForm.Items.Item("Sub3").DisplayDesc = true;

                // 순서
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("No").Specific, "Y");
                oForm.Items.Item("No").DisplayDesc = true;

                // Level
                oForm.Items.Item("Level").Specific.ValidValues.Add("0", "0");
                oForm.Items.Item("Level").Specific.ValidValues.Add("1", "1");
                oForm.Items.Item("Level").Specific.ValidValues.Add("2", "2");
                oForm.Items.Item("Level").DisplayDesc = true;

                // FatherId
                sQry = "select  distinct t.a,t.b   from (select distinct UniqueID as a , UniqueID as b from Authority_Folder union all select distinct FatherID as a , FatherID as b from Authority_Folder) t";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("FatherID").Specific, "Y");
                oForm.Items.Item("FatherID").DisplayDesc = true;

                // String
                oForm.DataSources.UserDataSources.Add("Strings", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("Strings").Specific.DataBind.SetBound(true, "", "Strings");

                // UniqueID
                oForm.DataSources.UserDataSources.Add("UniqueID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("UniqueID").Specific.DataBind.SetBound(true, "", "UniqueID");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY999_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY999_EnableMenus
        /// </summary>
        private void PH_PY999_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false);                // 제거
                oForm.EnableMenu("1284", false);                // 취소
                oForm.EnableMenu("1293", false);                // 행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY999_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY999_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if ((oForm.Items.Item("pFSGubun").Specific.VALUE == "F"))
                {
                    oForm.Items.Item("pUserID").Enabled = false;
                    oForm.Items.Item("CPUserID").Enabled = false;
                    oForm.Items.Item("pGubun").Enabled = true;
                    oForm.Items.Item("Modual").Enabled = true;
                    oForm.Items.Item("Sub1").Enabled = true;
                    oForm.Items.Item("Sub2").Enabled = true;
                    oForm.Items.Item("Sub3").Enabled = true;
                    oForm.Items.Item("No").Enabled = true;
                    oForm.Items.Item("Level").Enabled = true;
                    oForm.Items.Item("Strings").Enabled = true;
                    oForm.Items.Item("FatherID").Enabled = true;
                    oForm.Items.Item("Position").Enabled = true;
                    oForm.Items.Item("UserID").Enabled = false;
                    oForm.Items.Item("Sequence").Enabled = false;
                    oForm.Items.Item("UniqueID").Enabled = false;

                    oForm.Items.Item("Btn_Find").Enabled = true;
                    oForm.Items.Item("Bt_Copy").Enabled = false;

                }

                if ((oForm.Items.Item("pFSGubun").Specific.VALUE == "S"))
                {
                    oForm.Items.Item("pUserID").Enabled = true;
                    oForm.Items.Item("CPUserID").Enabled = false;
                    oForm.Items.Item("pGubun").Enabled = true;
                    oForm.Items.Item("Modual").Enabled = true;
                    oForm.Items.Item("Sub1").Enabled = true;
                    oForm.Items.Item("Sub2").Enabled = true;
                    oForm.Items.Item("Sub3").Enabled = true;
                    oForm.Items.Item("No").Enabled = true;
                    oForm.Items.Item("Level").Enabled = false;
                    oForm.Items.Item("Strings").Enabled = true;
                    oForm.Items.Item("FatherID").Enabled = true;
                    oForm.Items.Item("Position").Enabled = true;
                    oForm.Items.Item("UserID").Enabled = true;
                    oForm.Items.Item("Sequence").Enabled = false;
                    oForm.Items.Item("UniqueID").Enabled = true;

                    oForm.Items.Item("Btn_Find").Enabled = true;
                    oForm.Items.Item("Bt_Copy").Enabled = false;

                }
                if ((oForm.Items.Item("pFSGubun").Specific.VALUE == "C"))
                {
                    oForm.Items.Item("CPUserID").Enabled = true;
                    oForm.Items.Item("pUserID").Enabled = true;
                    oForm.Items.Item("pGubun").Enabled = false;
                    oForm.Items.Item("Modual").Enabled = false;
                    oForm.Items.Item("Sub1").Enabled = false;
                    oForm.Items.Item("Sub2").Enabled = false;
                    oForm.Items.Item("Sub3").Enabled = false;
                    oForm.Items.Item("No").Enabled = false;
                    oForm.Items.Item("Level").Enabled = false;
                    oForm.Items.Item("Strings").Enabled = false;
                    oForm.Items.Item("FatherID").Enabled = false;
                    oForm.Items.Item("Position").Enabled = false;
                    oForm.Items.Item("UserID").Enabled = false;
                    oForm.Items.Item("Sequence").Enabled = false;
                    oForm.Items.Item("UniqueID").Enabled = false;

                    oForm.Items.Item("Btn_Find").Enabled = false;
                    oForm.Items.Item("Bt_Copy").Enabled = true;
                }

                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    oForm.EnableMenu("1281", false);      // 문서찾기
                    oForm.EnableMenu("1282", true);       // 문서추가
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    oForm.EnableMenu("1281", false);      // 문서찾기
                    oForm.EnableMenu("1282", true);       // 문서추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    oForm.EnableMenu("1281", true);       // 문서찾기
                    oForm.EnableMenu("1282", true);       // 문서추가
                }
                // Key set
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPH_PY999_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    SubMain.Remove_Forms(oFormUniqueID01);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY999A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid1);
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
        /// Form Item Event
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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

                ////case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
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
                            PH_PY999_FormItemEnabled();
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        //Case "1293":
                        //  Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent);
                        case "1281": //문서찾기
                            PH_PY999_FormItemEnabled();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            break;
                        case "1293": // 행삭제
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
                    }

                    if (pVal.ItemUID == "Btn_Find")
                    {
                        PH_PY999_MTX01();
                    }

                    if (pVal.ItemUID == "Btn01")
                    {
                        PH_PY999_SAVE(pVal.Row);
                    }

                    if (pVal.ItemUID == "Btn_del")
                    {
                        PH_PY999_Delete();
                    }
                    if (pVal.ItemUID == "Bt_Copy")
                    {
                        PH_PY999_Copy();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.ItemUID)
                    {
                        case "1":
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY999_FormItemEnabled();
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY999_FormItemEnabled();
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY999_FormItemEnabled();
                                }
                            }
                            break;
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
                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
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
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        // 사업장(헤더)
                        switch (pVal.ItemUID)
                        {
                            case "Modual":
                            case "No":
                            case "Sub1":
                            case "Sub2":
                            case "Sub3":
                                if ((oForm.Items.Item("pFSGubun").Specific.VALUE == "F"))
                                {
                                    oForm.DataSources.UserDataSources.Item("UniqueId").Value = oForm.Items.Item("Modual").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub1").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub2").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub3").Specific.VALUE.ToString().Trim() + oForm.Items.Item("No").Specific.VALUE.ToString().Trim() + oForm.Items.Item("pFSGubun").Specific.VALUE.ToString().Trim();
                                    oForm.Items.Item("Sequence").Specific.VALUE = oForm.Items.Item("Modual").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub1").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub2").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub3").Specific.VALUE.ToString().Trim() + oForm.Items.Item("No").Specific.VALUE.ToString().Trim() + oForm.Items.Item("pFSGubun").Specific.VALUE.ToString().Trim();
                                }
                                oForm.Items.Item("Sequence").Specific.VALUE = oForm.Items.Item("Modual").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub1").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub2").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub3").Specific.VALUE.ToString().Trim() + oForm.Items.Item("No").Specific.VALUE.ToString().Trim() + oForm.Items.Item("pFSGubun").Specific.VALUE.ToString().Trim();
                                break;
                        }
                    }
                }
                PH_PY999_FormItemEnabled();
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Grid01":
                            if (pVal.Row >= 0)
                            {
                                switch (pVal.ItemUID)
                                {
                                    case "Grid01":
                                        PH_PY999_MTX02(pVal.ItemUID, pVal.Row, pVal.ColUID);
                                        break;
                                }
                            }
                            break;
                    }

                    switch (pVal.ItemUID)
                    {
                        case "Grid01":
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
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY999_MTX01
        /// </summary>
        private void PH_PY999_MTX01()
        {
            int iRow = 0;
            int ErrNum = 0;
            string sQry = string.Empty;
            string Param01 = string.Empty;
            string Param02 = string.Empty;
            string Param03 = string.Empty;

            try
            {
                oForm.Freeze(true);

                Param01 = oForm.Items.Item("pGubun").Specific.VALUE.ToString().Trim();
                Param02 = oForm.Items.Item("pFSGubun").Specific.VALUE.ToString().Trim();
                Param03 = oForm.Items.Item("pUserID").Specific.VALUE.ToString().Trim();

                if (string.IsNullOrEmpty(Strings.Trim(Param01)))
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                if (Param02 == "S")
                {
                    if (string.IsNullOrEmpty(Strings.Trim(Param03)))
                    {
                        ErrNum = 2;
                        throw new Exception();
                    }
                }

                sQry = "EXEC PH_PY999_01 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "'";
                oDS_PH_PY999A.ExecuteQuery(sQry);
                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("구분이 없습니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("USERID가 없습니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY999_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY999_MTX02
        /// </summary>
        private void PH_PY999_MTX02(string oUID, int oRow = 0, string oCol = "")
        {
            int sRow = 0;
            int ErrNum = 0;
            string sQry = string.Empty;
            string Param01 = string.Empty;
            string Param02 = string.Empty;
            string Param03 = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                sRow = oRow;
                Param01 = oForm.Items.Item("pFSGubun").Specific.VALUE.ToString().Trim();
                Param02 = oDS_PH_PY999A.Columns.Item("UniqueID").Cells.Item(oRow).Value;
                if (Param01 != "F")
                {
                    Param03 = oDS_PH_PY999A.Columns.Item("UserID").Cells.Item(oRow).Value;
                }
                sQry = "EXEC PH_PY999_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "'";
                oRecordSet.DoQuery(sQry);

                if ((oRecordSet.RecordCount == 0))
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                // Screen일때 UserID를 가져옴.
                if (Param01 != "F")
                {
                   // oForm.Items.Item("UserID").Specific.Select(oRecordSet.Fields.Item("UserID").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("UserID").Specific.value = oRecordSet.Fields.Item("UserID").Value;
                }
                // Folder일때 Level을 가져옴.
                if (Param01 != "S")
                {
                    //oForm.Items.Item("Level").Specific.Select(oRecordSet.Fields.Item("Level").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("Level").Specific.Select(oRecordSet.Fields.Item("Level").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                }

                //공통 S
                oForm.Items.Item("Modual").Specific.Select(oRecordSet.Fields.Item("Modual").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("Sub1").Specific.Select(oRecordSet.Fields.Item("Sub1").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("Sub2").Specific.Select(oRecordSet.Fields.Item("Sub2").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("Sub3").Specific.Select(oRecordSet.Fields.Item("Sub3").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("No").Specific.Select(oRecordSet.Fields.Item("No").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("Position").Specific.Select(oRecordSet.Fields.Item("Position").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("FatherID").Specific.Select(oRecordSet.Fields.Item("FatherID").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.DataSources.UserDataSources.Item("Strings").Value = oRecordSet.Fields.Item("Strings").Value;
                oForm.DataSources.UserDataSources.Item("UniqueId").Value = oRecordSet.Fields.Item("UniqueID").Value;
                oForm.Items.Item("Sequence").Specific.VALUE = oRecordSet.Fields.Item("Sequence").Value;
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY999_MTX02_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY999_SAVE
        /// </summary>
        private void PH_PY999_SAVE(int oRow = 0)
        {
            // 데이타 저장
            int ErrNum = 0;
            string sQry = string.Empty;
            string pGubun = string.Empty;
            string pFSGubun = string.Empty;
            string pUserID = string.Empty;
            string Modual = string.Empty;
            string Sub1 = string.Empty;
            string Sub2 = string.Empty;
            string Sub3 = string.Empty;
            string UserID = string.Empty;
            string No = string.Empty;
            string Level = string.Empty;
            string FatherID = string.Empty;
            string Position = string.Empty;
            string Strings_Renamed = string.Empty;
            string UniqueID = string.Empty;
            string pUniqueID = string.Empty;
            string Sequence = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                sQry = "select UniqueID, Seq  from Authority_Screen where UserID ='";
                sQry = sQry + oForm.Items.Item("UserID").Specific.VALUE + "'";
                oRecordSet.DoQuery(sQry);

                for (int i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (oRecordSet.Fields.Item(0).Value == oForm.Items.Item("UniqueID").Specific.VALUE.ToString().Trim())
                    {
                        ErrNum = 1;
                        throw new Exception();
                    }
                    if (oRecordSet.Fields.Item(1).Value  == oForm.Items.Item("Sequence").Specific.VALUE.ToString().Trim())
                    {
                        ErrNum = 1;
                        throw new Exception();
                    }
                    oRecordSet.MoveNext();
                }

                if (PSH_Globals.SBO_Application.MessageBox("데이터 입력하시겠습니까?", 2, "Yes", "No") == 2)
                {
                    ErrNum = 2;
                    throw new Exception();
                }

                pGubun = oForm.Items.Item("pGubun").Specific.VALUE.ToString().Trim();
                pFSGubun = oForm.Items.Item("pFSGubun").Specific.VALUE.ToString().Trim();
                pUserID = oForm.Items.Item("pUserID").Specific.VALUE.ToString().Trim();
                Modual = oForm.Items.Item("Modual").Specific.VALUE.ToString().Trim();
                Sub1 = oForm.Items.Item("Sub1").Specific.VALUE.ToString().Trim();
                Sub2 = oForm.Items.Item("Sub2").Specific.VALUE.ToString().Trim();
                Sub3 = oForm.Items.Item("Sub3").Specific.VALUE.ToString().Trim();
                UserID = oForm.Items.Item("UserID").Specific.VALUE.ToString().Trim();
                No = oForm.Items.Item("No").Specific.VALUE.ToString().Trim();
                Level = oForm.Items.Item("Level").Specific.VALUE.ToString().Trim();
                FatherID = oForm.Items.Item("FatherID").Specific.VALUE.ToString().Trim();
                Position = oForm.Items.Item("Position").Specific.VALUE.ToString().Trim();
                Strings_Renamed = oForm.Items.Item("Strings").Specific.VALUE.ToString().Trim();
                UniqueID = oForm.Items.Item("UniqueID").Specific.VALUE.ToString().Trim();
                Sequence = oForm.Items.Item("Sequence").Specific.VALUE.ToString().Trim();
                pUniqueID = oForm.Items.Item("UniqueID").Specific.VALUE.ToString().Trim();

                sQry = "EXEC PH_PY999_03 '" + pFSGubun + "', '" + pUserID + "', '" + pUniqueID + "', '";
                sQry = sQry + Sequence + "', '" + UniqueID + "', '" + UserID + "', '" + FatherID + "', '" + Strings_Renamed + "', '";
                sQry = sQry + Position + "', '" + Level + "', '" + No + "', '" + pGubun + "', '";
                sQry = sQry + PSH_Globals.oCompany.UserName + "'";
                oDS_PH_PY999A.ExecuteQuery(sQry);

                PSH_Globals.SBO_Application.StatusBar.SetText("입력완료", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                PH_PY999_MTX01();
            }
            catch (Exception ex)
            { 
                 if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("이미 저장된 값이 있습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("입력되지 않았습니다..", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY999_SAVE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY999_Delete
        /// </summary>
        private void PH_PY999_Delete()
        {
            // 데이타 삭제
            string sQry = string.Empty;
            int ErrNum = 0;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                
                if (PSH_Globals.SBO_Application.MessageBox("삭제하시겠습니까?", 2, "Yes", "No") == 2)
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                if ((oForm.Items.Item("pFSGubun").Specific.VALUE == "F"))
                {
                    sQry = "        delete";
                    sQry = sQry + " from    Authority_Folder";
                    sQry = sQry + " where   UniqueID = '" + oForm.Items.Item("UniqueID").Specific.VALUE + "'";
                }
                else
                {
                    sQry = "        delete";
                    sQry = sQry + " from    Authority_Screen";
                    sQry = sQry + " where   UniqueID = '" + oForm.Items.Item("UniqueID").Specific.VALUE + "'";
                    sQry = sQry + "         AND UserID = '" + oForm.Items.Item("pUserID").Specific.VALUE + "'";
                }
                oRecordSet.DoQuery(sQry);
                PH_PY999_MTX01();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("취소되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY999_Delete_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY999_Copy
        /// </summary>
        private void PH_PY999_Copy()
        {
            // 데이타 삭제
            string sQry = string.Empty;
            string pUserID = string.Empty;
            string CPUserID = string.Empty;
            int ErrNo = 0;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            pUserID = oForm.Items.Item("pUserID").Specific.VALUE.ToString().Trim();
            CPUserID = oForm.Items.Item("CPUserID").Specific.VALUE.ToString().Trim();
            try
            {
                oForm.Freeze(true);

                sQry = "select count(1) from Authority_Screen where UserID ='" + CPUserID + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount <= 0)
                {
                    ErrNo = 1;
                    throw new Exception(); 
                }
                if(pUserID == "" || CPUserID =="")
                {
                    ErrNo = 2;
                    throw new Exception();
                }
                else
                {
                    sQry = "delete from Authority_Screen where UniqueID ='" + oForm.Items.Item("UniqueID").Specific.VALUE + "'";
                }

                sQry = "Insert into Authority_Screen";
                sQry = sQry + " select '" + CPUserID + "'";
                sQry = sQry + ", FatherID";
                sQry = sQry + ", String";
                sQry = sQry + ", UniqueID";
                sQry = sQry + ", Position";
                sQry = sQry + ", Type";
                sQry = sQry + ", Seq";
                sQry = sQry + ", Gubun";
                sQry = sQry + ", 'Y'";
                sQry = sQry + ", GETDATE()";
                sQry = sQry + ", '" + PSH_Globals.oCompany.UserName  + "'";
                sQry = sQry + "  from Authority_Screen";
                sQry = sQry + "  where UserID ='" + pUserID  + "'";

                oRecordSet.DoQuery(sQry);
            }
            catch (Exception ex)
            {
                if (ErrNo == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("복제 계정에 권한이 있습니다. 복제 하고자 모든 권한을 삭제하세요.");
                }
                if (ErrNo == 2)
                {
                    PSH_Globals.SBO_Application.MessageBox("대상 ID와 복제ID는 필수입니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY999_Copy_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
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
// // ERROR: Not supported in C#: OptionDeclaration
//namespace MDC_HR_Addon
//{
//	internal class PH_PY999
//	{
//////********************************************************************************
//////  File           : PH_PY999.cls
//////  Module         :
//////  Desc           :
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		public SAPbouiCOM.Grid oGrid1;
//		public SAPbouiCOM.DataTable oDS_PH_PY999A;


//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY999.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY999_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY999");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

//			//oForm.PaneLevel = 1
//			oForm.Freeze(true);
//			PH_PY999_CreateItems();
//			PH_PY999_FormItemEnabled();
//			PH_PY999_EnableMenus();

//			oForm.Update();
//			oForm.Freeze(false);

//			oForm.Visible = true;
//			//UPGRADE_NOTE: Object oXmlDoc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			return;
//			LoadForm_Error:

//			oForm.Update();
//			oForm.Freeze(false);
//			//UPGRADE_NOTE: Object oXmlDoc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			//UPGRADE_NOTE: Object oForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oForm = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private bool PH_PY999_CreateItems()
//		{
//			bool functionReturnValue = false;

//			string sQry = null;
//			int i = 0;
//			string CLTCOD = null;

//			SAPbouiCOM.CheckBox oCheck = null;
//			SAPbouiCOM.EditText oEdit = null;
//			SAPbouiCOM.ComboBox oCombo = null;
//			//Dim oColumn     As SAPbouiCOM.Column
//			//Dim oColumns    As SAPbouiCOM.Columns
//			//Dim optBtn      As SAPbouiCOM.OptionBtn

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oGrid1 = oForm.Items.Item("Grid01").Specific;

//			oForm.DataSources.DataTables.Add("PH_PY999");

//			oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY999");
//			oDS_PH_PY999A = oForm.DataSources.DataTables.Item("PH_PY999");


//			////----------------------------------------------------------------------------------------------
//			//// 기본사항
//			////----------------------------------------------------------------------------------------------

//			////구분
//			oCombo = oForm.Items.Item("pGubun").Specific;
//			oCombo.ValidValues.Add("B", "기본");
//			oCombo.ValidValues.Add("H", "인사");
//			oForm.Items.Item("pGubun").DisplayDesc = true;

//			////폴더/화면구분
//			oCombo = oForm.Items.Item("pFSGubun").Specific;
//			oCombo.ValidValues.Add("F", "폴더");
//			oCombo.ValidValues.Add("S", "화면");
//			oForm.Items.Item("pFSGubun").DisplayDesc = true;

//			////pUserID
//			oCombo = oForm.Items.Item("pUserID").Specific;
//			sQry = "select USER_CODE,U_NAME from ousr";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oForm.Items.Item("pUserID").DisplayDesc = true;

//			////순서
//			oCombo = oForm.Items.Item("Modual").Specific;
//			sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='A006'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oForm.Items.Item("Modual").DisplayDesc = true;

//			////순서
//			oCombo = oForm.Items.Item("Position").Specific;
//			sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='A005'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oForm.Items.Item("Position").DisplayDesc = true;

//			////Sub1
//			oCombo = oForm.Items.Item("Sub1").Specific;
//			sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oForm.Items.Item("Sub1").DisplayDesc = true;

//			////Sub2
//			oCombo = oForm.Items.Item("Sub2").Specific;
//			sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oForm.Items.Item("Sub2").DisplayDesc = true;

//			////Sub3
//			oCombo = oForm.Items.Item("Sub3").Specific;
//			sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oForm.Items.Item("Sub3").DisplayDesc = true;

//			////순서
//			oCombo = oForm.Items.Item("No").Specific;
//			sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oForm.Items.Item("No").DisplayDesc = true;

//			////Level
//			oCombo = oForm.Items.Item("Level").Specific;
//			oCombo.ValidValues.Add("0", "0");
//			oCombo.ValidValues.Add("1", "1");
//			oCombo.ValidValues.Add("2", "2");
//			oForm.Items.Item("Level").DisplayDesc = true;

//			////UserID
//			oCombo = oForm.Items.Item("UserID").Specific;
//			sQry = "select USER_CODE,U_NAME from ousr";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oForm.Items.Item("UserID").DisplayDesc = true;

//			////FatherId
//			oCombo = oForm.Items.Item("FatherID").Specific;
//			sQry = "select  distinct t.a,t.b   from (select distinct UniqueID as a , UniqueID as b from Authority_Folder union all select distinct FatherID as a , FatherID as b from Authority_Folder) t";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oForm.Items.Item("FatherID").DisplayDesc = true;

//			////String
//			oForm.DataSources.UserDataSources.Add("Strings", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.DataBind. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Strings").Specific.DataBind.SetBound(true, "", "Strings");

//			////UniqueID
//			oForm.DataSources.UserDataSources.Add("UniqueID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.DataBind. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("UniqueID").Specific.DataBind.SetBound(true, "", "UniqueID");

//			oForm.Update();

//			//UPGRADE_NOTE: Object oCheck may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: Object oEdit may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: Object oCombo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//Set oColumn = Nothing
//			//Set oColumns = Nothing
//			//Set optBtn = Nothing
//			//UPGRADE_NOTE: Object oRecordSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return functionReturnValue;
//			PH_PY999_CreateItems_Error:

//			//UPGRADE_NOTE: Object oCheck may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: Object oEdit may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: Object oCombo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//Set oColumn = Nothing
//			//Set oColumns = Nothing
//			//Set optBtn = Nothing
//			//UPGRADE_NOTE: Object oRecordSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY999_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY999_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.EnableMenu("1283", false);
//			////제거
//			oForm.EnableMenu("1284", false);
//			////취소
//			oForm.EnableMenu("1293", false);
//			////행삭제

//			return;
//			PH_PY999_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY999_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY999_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY999_FormItemEnabled();
//				//        Call PH_PY999_AddMatrixRow
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY999_FormItemEnabled();
//				//oForm.Items("Code").Specific.VALUE = oFromDocEntry01
//				//oForm.Items("1").CLICK ct_Regular
//			}
//			return;
//			PH_PY999_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY999_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY999_FormItemEnabled()
//		{
//			SAPbouiCOM.ComboBox oCombo = null;
//			string sQry = null;
//			int i = 0;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze(true);

//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items(pFSGubun).Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if ((oForm.Items.Item("pFSGubun").Specific.VALUE == "F")) {
//				oForm.Items.Item("pUserID").Enabled = false;
//				oForm.Items.Item("UserID").Enabled = false;
//				oForm.Items.Item("Level").Enabled = true;
//				oForm.Items.Item("UniqueID").Enabled = false;
//			}

//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items(pFSGubun).Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if ((oForm.Items.Item("pFSGubun").Specific.VALUE == "S")) {
//				oForm.Items.Item("Level").Enabled = false;
//				oForm.Items.Item("pUserID").Enabled = true;
//				oForm.Items.Item("UserID").Enabled = true;
//				oForm.Items.Item("UniqueID").Enabled = true;
//			}


//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {


//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가


//				//// 접속자에 따른 권한별 사업장 콤보박스세팅

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅

//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가


//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			}

//			////Key set


//			//UPGRADE_NOTE: Object oCombo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY999_FormItemEnabled_Error:

//			//UPGRADE_NOTE: Object oCombo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY999_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			string sQry = null;
//			int i = 0;
//			string temp = null;

//			//Dim oCombo      As SAPbouiCOM.ComboBox
//			// Dim oColumn     As SAPbouiCOM.Column
//			// Dim oColumns     As SAPbouiCOM.Columns
//			//Dim oRecordSet  As SAPbobsCOM.Recordset

//			 // ERROR: Not supported in C#: OnErrorStatement


//			//Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1

//					if (pval.BeforeAction == true) {
//						if (pval.ItemUID == "1") {
//							//                    If PH_PY999_DataValidCheck = False Then
//							//                        BubbleEvent = False
//							//                    End If
//						}

//						if (pval.ItemUID == "Btn_Find") {
//							PH_PY999_MTX01();
//						}

//						if (pval.ItemUID == "Btn01") {
//							PH_PY999_SAVE(ref pval.Row);
//						}

//						if (pval.ItemUID == "Btn_del") {
//							PH_PY999_Delete();
//						}

//					} else if (pval.BeforeAction == false) {
//						switch (pval.ItemUID) {
//							case "1":
//								if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//									if (pval.ActionSuccess == true) {
//										PH_PY999_FormItemEnabled();
//									}
//								} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//									if (pval.ActionSuccess == true) {
//										PH_PY999_FormItemEnabled();
//									}
//								} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//									if (pval.ActionSuccess == true) {
//										PH_PY999_FormItemEnabled();
//									}
//								}
//								break;
//							//
//						}
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2
//					if (pval.BeforeAction == true) {
//						if (pval.CharPressed == 9) {
//							//                    If pval.ItemUID = "MSTCOD" Then
//							//                        If oForm.Items("MSTCOD").Specific.VALUE = "" Then
//							//                            Sbo_Application.ActivateMenuItem ("7425")
//							//                                BubbleEvent = False
//							//                        End If
//							//                    End If
//						}
//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					switch (pval.ItemUID) {
//						case "Mat01":
//							if (pval.Row > 0) {
//								oLastItemUID = pval.ItemUID;
//								oLastColUID = pval.ColUID;
//								oLastColRow = pval.Row;
//							}
//							break;
//						default:
//							oLastItemUID = pval.ItemUID;
//							oLastColUID = "";
//							oLastColRow = 0;
//							break;
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//					////4
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//					////5
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemChanged == true) {
//							//                    //사업장(헤더)
//							switch (pval.ItemUID) {
//								case "Modual":
//								case "No":
//								case "Sub1":
//								case "Sub2":
//								case "Sub3":
//									//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items(pFSGubun).Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if ((oForm.Items.Item("pFSGubun").Specific.VALUE == "F")) {
//										//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.DataSources.UserDataSources.Item("UniqueId").Value = Strings.Trim(oForm.Items.Item("Modual").Specific.VALUE) + Strings.Trim(oForm.Items.Item("Sub1").Specific.VALUE) + Strings.Trim(oForm.Items.Item("Sub2").Specific.VALUE) + Strings.Trim(oForm.Items.Item("Sub3").Specific.VALUE) + Strings.Trim(oForm.Items.Item("No").Specific.VALUE) + Strings.Trim(oForm.Items.Item("pFSGubun").Specific.VALUE);
//										//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items(Sequence).Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("Sequence").Specific.VALUE = Strings.Trim(oForm.Items.Item("Modual").Specific.VALUE) + Strings.Trim(oForm.Items.Item("Sub1").Specific.VALUE) + Strings.Trim(oForm.Items.Item("Sub2").Specific.VALUE) + Strings.Trim(oForm.Items.Item("Sub3").Specific.VALUE) + Strings.Trim(oForm.Items.Item("No").Specific.VALUE) + Strings.Trim(oForm.Items.Item("pFSGubun").Specific.VALUE);
//									}
//									//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items(Sequence).Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("Sequence").Specific.VALUE = Strings.Trim(oForm.Items.Item("Modual").Specific.VALUE) + Strings.Trim(oForm.Items.Item("Sub1").Specific.VALUE) + Strings.Trim(oForm.Items.Item("Sub2").Specific.VALUE) + Strings.Trim(oForm.Items.Item("Sub3").Specific.VALUE) + Strings.Trim(oForm.Items.Item("No").Specific.VALUE) + Strings.Trim(oForm.Items.Item("pFSGubun").Specific.VALUE);
//									break;
//							}
//						}
//					}
//					PH_PY999_FormItemEnabled();
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {
//						switch (pval.ItemUID) {
//							case "Grid01":
//								if (pval.Row >= 0) {
//									switch (pval.ItemUID) {
//										case "Grid01":
//											PH_PY999_MTX02(pval.ItemUID, ref pval.Row, ref pval.ColUID);
//											break;
//									}

//								}
//								break;
//						}

//						switch (pval.ItemUID) {
//							case "Grid01":
//								if (pval.Row > 0) {
//									oLastItemUID = pval.ItemUID;
//									oLastColUID = pval.ColUID;
//									oLastColRow = pval.Row;
//								}
//								break;
//							default:
//								oLastItemUID = pval.ItemUID;
//								oLastColUID = "";
//								oLastColRow = 0;
//								break;
//						}
//					} else if (pval.BeforeAction == false) {

//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//					////7
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {
//					} else {

//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//					////8
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
//					////9
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					////10
//					//            Call oForm.Freeze(True)
//					if (pval.BeforeAction == true) {
//						if (pval.ItemChanged == true) {

//						}

//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemChanged == true) {
//							switch (pval.ItemUID) {

//							}

//						}
//					}
//					break;
//				//            Call oForm.Freeze(False)
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					////11
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						//                oMat1.LoadFromDataSource
//						//                Call PH_PY999_AddMatrixRow

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
//					////12
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
//					////16
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					////17
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: Object oForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm = null;
//						//UPGRADE_NOTE: Object oDS_PH_PY999A may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY999A = null;

//						//                Set oMat1 = Nothing
//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//					////18
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//					////19
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
//					////20
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//					////21
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
//					////22
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
//					////23
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//					////27
//					if (pval.BeforeAction == true) {

//					} else if (pval.Before_Action == false) {
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
//					////37
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
//					////38
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_Drag:
//					////39
//					break;


//			}

//			//Set oCombo = Nothing
//			//Set oRecordSet = Nothing

//			return;
//			PH_PY999_FormItemEvent_Exit:
//			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			//Set oRecordSet = Nothing
//			oForm.Freeze(false);

//			return;
//			Raise_FormItemEvent_Error:

//			oForm.Freeze((false));
//			//Set oCombo = Nothing
//			//Set oRecordSet = Nothing
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{
//			int i = 0;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			oForm.Freeze(true);

//			if ((pval.BeforeAction == true)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						if (MDC_Globals.Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2) {
//							BubbleEvent = false;
//							return;
//						}
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					case "1293":
//						break;
//					case "1281":
//						break;
//					case "1282":
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY999_FormItemEnabled();
//						break;
//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						break;
//					//Call PH_PY999_FormItemEnabled
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						//Call PH_PY999_FormItemEnabled
//						oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1282":
//						////문서추가
//						break;
//					//Call PH_PY999_FormItemEnabled
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						break;
//					//Call PH_PY999_FormItemEnabled
//					case "1293":
//						//// 행삭제
//						break;

//				}
//			}
//			oForm.Freeze(false);
//			return;
//			Raise_FormMenuEvent_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{
//			int i = 0;
//			string sQry = null;
//			SAPbouiCOM.ComboBox oCombo = null;

//			SAPbobsCOM.Recordset oRecordSet = null;


//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			if ((BusinessObjectInfo.BeforeAction == false)) {
//				switch (BusinessObjectInfo.EventType) {
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//						////33
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//						////34
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//						////35
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//						////36
//						break;
//				}

//			}
//			//UPGRADE_NOTE: Object oCombo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: Object oRecordSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Raise_FormDataEvent_Error:

//			//UPGRADE_NOTE: Object oCombo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: Object oRecordSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//		}

//		public void PH_PY999_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: Couldn't resolve default property of object MDC_GetData.Get_ReData(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY999'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//			} else {
//				//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//			}
//			return;
//			PH_PY999_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY999_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY999_MTX01()
//		{

//			////그리드에 데이터 로드

//			int i = 0;
//			string sQry = null;
//			int iRow = 0;

//			string Param01 = null;
//			string Param02 = null;
//			string Param03 = null;


//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param01 = Strings.Trim(oForm.Items.Item("pGubun").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param02 = Strings.Trim(oForm.Items.Item("pFSGubun").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param03 = Strings.Trim(oForm.Items.Item("pUserID").Specific.VALUE);

//			if (string.IsNullOrEmpty(Strings.Trim(Param01))) {
//				MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY999_MTX01_Exit;
//			}


//			sQry = "EXEC PH_PY999_01 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "'";

//			oDS_PH_PY999A.ExecuteQuery(sQry);



//			iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;

//			oForm.Update();

//			//UPGRADE_NOTE: Object oRecordSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY999_MTX01_Exit:
//			//UPGRADE_NOTE: Object oRecordSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY999_MTX01_Error:
//			//UPGRADE_NOTE: Object oRecordSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY999_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}
//		private void PH_PY999_MTX02(string oUID, ref int oRow = 0, ref string oCol = "")
//		{

//			////그리드 자료를 head에 로드

//			int i = 0;
//			string sQry = null;
//			int sRow = 0;

//			string Param01 = null;
//			string Param02 = null;
//			string Param03 = null;

//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sRow = oRow;

//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param01 = Strings.Trim(oForm.Items.Item("pFSGubun").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oDS_PH_PY999A.Columns.Item().Cells().VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param02 = oDS_PH_PY999A.Columns.Item("UniqueID").Cells.Item(oRow).Value;
//			if (Param01 != "F") {
//				//UPGRADE_WARNING: Couldn't resolve default property of object oDS_PH_PY999A.Columns.Item().Cells().VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				Param03 = oDS_PH_PY999A.Columns.Item("UserID").Cells.Item(oRow).Value;
//			}

//			sQry = "EXEC PH_PY999_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "'";
//			oRecordSet.DoQuery(sQry);

//			if ((oRecordSet.RecordCount == 0)) {

//				MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
//				goto PH_PY999_MTX02_Exit;
//			}

//			// Screen일때 UserID를 가져옴.
//			if (Param01 != "F") {
//				oCombo = oForm.Items.Item("UserID").Specific;
//				oCombo.Select(oRecordSet.Fields.Item("UserID").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
//			}
//			// Folder일때 Level을 가져옴.
//			if (Param01 != "S") {
//				oCombo = oForm.Items.Item("Level").Specific;
//				oCombo.Select(oRecordSet.Fields.Item("Level").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
//			}

//			//공통 S
//			oCombo = oForm.Items.Item("Modual").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("Modual").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

//			oCombo = oForm.Items.Item("Sub1").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("Sub1").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

//			oCombo = oForm.Items.Item("Sub2").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("Sub2").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

//			oCombo = oForm.Items.Item("Sub3").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("Sub3").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

//			oCombo = oForm.Items.Item("No").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("No").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

//			oCombo = oForm.Items.Item("Position").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("Position").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

//			oCombo = oForm.Items.Item("FatherID").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("FatherID").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

//			//UPGRADE_WARNING: Couldn't resolve default property of object oRecordSet.Fields().VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("Strings").Value = oRecordSet.Fields.Item("Strings").Value;
//			//UPGRADE_WARNING: Couldn't resolve default property of object oRecordSet.Fields().VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("UniqueId").Value = oRecordSet.Fields.Item("UniqueID").Value;
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items(Sequence).Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: Couldn't resolve default property of object oRecordSet.Fields().VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Sequence").Specific.VALUE = oRecordSet.Fields.Item("Sequence").Value;

//			//공통E

//			//UPGRADE_NOTE: Object oRecordSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY999_MTX02_Exit:
//			//UPGRADE_NOTE: Object oRecordSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);

//			return;
//			PH_PY999_MTX02_Error:
//			//UPGRADE_NOTE: Object oRecordSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY999_MTX02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}
//		private void PH_PY999_SAVE(ref int oRow = 0)
//		{

//			////데이타 저장

//			int i = 0;
//			string sQry = null;

//			string pGubun = null;
//			string pFSGubun = null;
//			string pUserID = null;
//			string Modual = null;
//			string Sub1 = null;
//			string Sub2 = null;
//			string Sub3 = null;
//			string UserID = null;
//			string No = null;
//			string Level = null;
//			string FatherID = null;
//			string Position = null;
//			//UPGRADE_NOTE: Strings was upgraded to Strings_Renamed. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
//			string Strings_Renamed = null;
//			string UniqueID = null;
//			string pUniqueID = null;
//			string Sequence = null;

//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			pGubun = Strings.Trim(oForm.Items.Item("pGubun").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			pFSGubun = Strings.Trim(oForm.Items.Item("pFSGubun").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			pUserID = Strings.Trim(oForm.Items.Item("pUserID").Specific.VALUE);

//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Modual = Strings.Trim(oForm.Items.Item("Modual").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Sub1 = Strings.Trim(oForm.Items.Item("Sub1").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Sub2 = Strings.Trim(oForm.Items.Item("Sub2").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Sub3 = Strings.Trim(oForm.Items.Item("Sub3").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			UserID = Strings.Trim(oForm.Items.Item("UserID").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			No = Strings.Trim(oForm.Items.Item("No").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Level = Strings.Trim(oForm.Items.Item("Level").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			FatherID = Strings.Trim(oForm.Items.Item("FatherID").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Position = Strings.Trim(oForm.Items.Item("Position").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Strings_Renamed = Strings.Trim(oForm.Items.Item("Strings").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			UniqueID = Strings.Trim(oForm.Items.Item("UniqueID").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Sequence = Strings.Trim(oForm.Items.Item("Sequence").Specific.VALUE);
//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			pUniqueID = Strings.Trim(oForm.Items.Item("UniqueID").Specific.VALUE);

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "EXEC PH_PY999_03 '" + pFSGubun + "', '" + pUserID + "', '" + pUniqueID + "', '";
//			sQry = sQry + Sequence + "', '" + UniqueID + "', '" + UserID + "', '" + FatherID + "', '" + Strings_Renamed + "', '";
//			sQry = sQry + Position + "', '" + Level + "', '" + No + "', '" + pGubun + "', '";
//			sQry = sQry + MDC_Globals.oCompany.UserName + "'";
//			oDS_PH_PY999A.ExecuteQuery(sQry);

//			//UPGRADE_NOTE: Object oRecordSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);

//			PH_PY999_MTX01();

//			return;
//			PH_PY999_SAVE_Exit:

//			//UPGRADE_NOTE: Object oRecordSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);

//			return;
//			PH_PY999_SAVE_Error:
//			oForm.Freeze(false);

//			//UPGRADE_NOTE: Object oRecordSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY999_SAVE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		private void PH_PY999_Delete(ref int oRow = 0)
//		{
//			////선택된 자료 삭제

//			string pUniqueID = null;
//			string sQry = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze(true);

//			//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items(pFSGubun).Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if ((oForm.Items.Item("pFSGubun").Specific.VALUE == "F")) {
//				//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = "delete from Authority_Folder where UniqueID ='" + oForm.Items.Item("UniqueID").Specific.VALUE + "'";
//			} else {
//				//UPGRADE_WARNING: Couldn't resolve default property of object oForm.Items().Specific.VALUE. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = "delete from Authority_Screen where UniqueID ='" + oForm.Items.Item("UniqueID").Specific.VALUE + "'";
//			}

//			oRecordSet.DoQuery(sQry);

//			oForm.Freeze(false);

//			PH_PY999_MTX01();

//			return;
//			PH_PY999_Delete_Exit:

//			//UPGRADE_NOTE: Object oRecordSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);

//			return;
//			PH_PY999_Delete_Error:
//			oForm.Freeze(false);

//			//UPGRADE_NOTE: Object oRecordSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY999_Delete_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

////
////Private Sub PH_PY999_TitleSetting(iRow As Long)
////
////End Sub
////
////
////
////
//	}
//}
