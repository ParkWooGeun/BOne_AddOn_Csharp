using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 일일근태처리
    /// </summary>
    internal class PH_PY010 : PSH_BaseClass
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY010.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY010_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY010");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY010_CreateItems();
                PH_PY010_FormItemEnabled();
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
        private void PH_PY010_CreateItems()
        {
            string sQry = string.Empty;
            string CLTCOD = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = null;

            string DocDateF = string.Empty;
            string DocDateT = string.Empty;
            string ImsiDate = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);

                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");

                oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");

                //// 사업장
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                sQry = "select Convert(char(8),GetDate(),112), Convert(char(8),DateAdd(dd, -1, GetDate()),112)";
                oRecordSet.DoQuery(sQry);

                DocDateT = oRecordSet.Fields.Item(0).Value;
                ImsiDate = oRecordSet.Fields.Item(1).Value;

                CLTCOD = dataHelpClass.Get_ReData("Branch", "USER_CODE", "OUSR", "'" + PSH_Globals.oCompany.UserName + "'","");


                sQry = "select Convert(char(8),Isnull(max(RToDate), '20130101') ,112) from ZPH_PY008 Where PosDate between DateAdd(mm, -1, '" + DocDateT + "') and '" + DocDateT + "' and CLTCOD = '" + CLTCOD + "'";
                oRecordSet.DoQuery(sQry);

                DocDateF = oRecordSet.Fields.Item(0).Value;

                if (DocDateF == DocDateT)
                {
                    oForm.Items.Item("DocDateFr").Specific.VALUE = ImsiDate;
                    oForm.Items.Item("DocDateTo").Specific.VALUE = DocDateT;
                }
                else
                {
                    oForm.Items.Item("DocDateFr").Specific.VALUE = DocDateF;
                    oForm.Items.Item("DocDateTo").Specific.VALUE = DocDateT;
                }
             }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY010_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        public void PH_PY010_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY010_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:                    ////2
                    break;


                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                    ////3

                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                    ////4
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
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {

                        }
                        ////해야할일 작업
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
                                PH_PY010_FormItemEnabled();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY010_FormItemEnabled();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY010_FormItemEnabled();
                            }
                        }
                    }
                    if (pVal.ItemUID == "Btn_Proc")
                    {
                        PH_PY010_Proc();
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
        /// 근태처리 로직
        /// </summary>
        private void PH_PY010_Proc()
        {
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string DocDateFr = string.Empty;
            string DocDateTo = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.Trim();
                DocDateFr = oForm.Items.Item("DocDateFr").Specific.VALUE.Trim();
                DocDateTo = oForm.Items.Item("DocDateTo").Specific.VALUE.Trim();

                sQry = "PH_PY010_01 '" + CLTCOD + "','" + DocDateFr + "', '" + DocDateTo + "'";
                oRecordSet.DoQuery(sQry);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY010_Proc_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                PSH_Globals.SBO_Application.MessageBox("근태처리 완료되었습니다.");
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
//	internal class PH_PY010
//	{
//////********************************************************************************
//////  File           : PH_PY010.cls
//////  Module         : 근태관리 > 일근태처리
//////  Desc           : 일근태처리
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		public SAPbouiCOM.Matrix oMat1;

//		private SAPbouiCOM.DBDataSource oDS_PH_PY010A;


//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY010.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY010_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY010");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


//			oForm.Freeze(true);
//			PH_PY010_CreateItems();
//			PH_PY010_EnableMenus();
//			PH_PY010_SetDocument(oFromDocEntry01);
//			//    Call PH_PY010_FormResize

//			oForm.Update();
//			oForm.Freeze(false);

//			oForm.Visible = true;
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			return;
//			LoadForm_Error:

//			oForm.Update();
//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oForm = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//private bool PH_PY010_CreateItems()
//{
//    bool functionReturnValue = false;

//    string sQry = null;
//    int i = 0;
//    string CLTCOD = null;

//    SAPbouiCOM.CheckBox oCheck = null;
//    SAPbouiCOM.EditText oEdit = null;
//    SAPbouiCOM.ComboBox oCombo = null;
//    SAPbouiCOM.Column oColumn = null;
//    SAPbouiCOM.Columns oColumns = null;
//    SAPbouiCOM.OptionBtn optBtn = null;

//    SAPbobsCOM.Recordset oRecordSet = null;

//    string DocDateF = null;
//    string DocDateT = null;
//    string ImsiDate = null;



//    // ERROR: Not supported in C#: OnErrorStatement


//    oForm.Freeze(true);

//    oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//    oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
//    //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//    oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");

//    oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
//    //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//    oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");

//    //// 사업장
//    oCombo = oForm.Items.Item("CLTCOD").Specific;
//    oForm.Items.Item("CLTCOD").DisplayDesc = true;

//    sQry = "select Convert(char(8),GetDate(),112), Convert(char(8),DateAdd(dd, -1, GetDate()),112)";
//    oRecordSet.DoQuery(sQry);

//    //UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//    DocDateT = oRecordSet.Fields.Item(0).Value;
//    //UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//    ImsiDate = oRecordSet.Fields.Item(1).Value;

//    //UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//    CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" + MDC_Globals.oCompany.UserName + "'");


//    sQry = "select Convert(char(8),Isnull(max(RToDate), '20130101') ,112) from ZPH_PY008 Where PosDate between DateAdd(mm, -1, '" + DocDateT + "') and '" + DocDateT + "' and CLTCOD = '" + CLTCOD + "'";
//    oRecordSet.DoQuery(sQry);

//    //UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//    DocDateF = oRecordSet.Fields.Item(0).Value;

//    if (DocDateF == DocDateT)
//    {
//        //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        oForm.Items.Item("DocDateFr").Specific.VALUE = ImsiDate;
//        //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        oForm.Items.Item("DocDateTo").Specific.VALUE = DocDateT;
//    }
//    else
//    {
//        //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        oForm.Items.Item("DocDateFr").Specific.VALUE = DocDateF;
//        //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        oForm.Items.Item("DocDateTo").Specific.VALUE = DocDateT;
//    }








//    //UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oCheck = null;
//    //UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oEdit = null;
//    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oCombo = null;
//    //UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oColumn = null;
//    //UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oColumns = null;
//    //UPGRADE_NOTE: optBtn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    optBtn = null;
//    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oRecordSet = null;
//    oForm.Freeze(false);
//    return functionReturnValue;
//PH_PY010_CreateItems_Error:

//    //UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oCheck = null;
//    //UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oEdit = null;
//    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oCombo = null;
//    //UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oColumn = null;
//    //UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oColumns = null;
//    //UPGRADE_NOTE: optBtn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    optBtn = null;
//    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oRecordSet = null;
//    oForm.Freeze(false);
//    MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY010_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//    return functionReturnValue;
//}

//		private void PH_PY010_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement

//			oForm.EnableMenu("1281", false);
//			////찾기
//			oForm.EnableMenu("1282", false);
//			////신규
//			oForm.EnableMenu("1283", false);
//			////제거
//			oForm.EnableMenu("1284", false);
//			////취소
//			oForm.EnableMenu("1293", false);
//			////행삭제

//			return;
//			PH_PY010_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY010_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY010_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY010_FormItemEnabled();

//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY010_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY010_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY010_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//public void PH_PY010_FormItemEnabled()
//{
//    SAPbouiCOM.ComboBox oCombo = null;

//    // ERROR: Not supported in C#: OnErrorStatement



//    oForm.Freeze(true);
//    if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
//    {
//        //// 접속자에 따른 권한별 사업장 콤보박스세팅
//        MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//        oForm.EnableMenu("1281", false);
//        ////문서찾기
//        oForm.EnableMenu("1282", false);
//        ////문서추가

//    }
//    else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
//    {
//        //// 접속자에 따른 권한별 사업장 콤보박스세팅
//        MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//        oForm.EnableMenu("1281", false);
//        ////문서찾기
//        oForm.EnableMenu("1282", false);
//        ////문서추가
//    }
//    else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
//    {
//        //// 접속자에 따른 권한별 사업장 콤보박스세팅
//        MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);

//        oForm.EnableMenu("1281", false);
//        ////문서찾기
//        oForm.EnableMenu("1282", false);
//        ////문서추가

//    }
//    oForm.Freeze(false);
//    return;
//PH_PY010_FormItemEnabled_Error:

//    oForm.Freeze(false);
//    MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY010_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//}


//public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//{
//    string sQry = null;
//    int i = 0;
//    SAPbouiCOM.ComboBox oCombo = null;
//    SAPbobsCOM.Recordset oRecordSet = null;

//    // ERROR: Not supported in C#: OnErrorStatement


//    oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//    switch (pVal.EventType)
//    {
//        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//            ////1

//            if (pVal.BeforeAction == true)
//            {
//                if (pVal.ItemUID == "1")
//                {
//                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                    {

//                    }
//                    ////해야할일 작업
//                }

//            }
//            else if (pVal.BeforeAction == false)
//            {
//                if (pVal.ItemUID == "1")
//                {
//                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                    {
//                        if (pVal.ActionSuccess == true)
//                        {
//                            PH_PY010_FormItemEnabled();
//                        }
//                    }
//                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                    {
//                        if (pVal.ActionSuccess == true)
//                        {
//                            PH_PY010_FormItemEnabled();
//                        }
//                    }
//                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//                    {
//                        if (pVal.ActionSuccess == true)
//                        {
//                            PH_PY010_FormItemEnabled();
//                        }
//                    }
//                }
//                if (pVal.ItemUID == "Btn_Proc")
//                {
//                    PH_PY010_Proc();
//                }

//            }
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//            ////2
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//            ////3
//            switch (pVal.ItemUID)
//            {
//                case "Mat01":
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
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//            ////4
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//            ////5
//            oForm.Freeze(true);
//            if (pVal.BeforeAction == true)
//            {

//            }
//            else if (pVal.BeforeAction == false)
//            {
//                if (pVal.ItemChanged == true)
//                {

//                }
//            }
//            oForm.Freeze(false);
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_CLICK:
//            ////6
//            if (pVal.BeforeAction == true)
//            {
//                switch (pVal.ItemUID)
//                {
//                    case "Mat01":
//                        if (pVal.Row > 0)
//                        {
//                            oMat1.SelectRow(pVal.Row, true, false);
//                        }
//                        break;
//                }

//                switch (pVal.ItemUID)
//                {
//                    case "Mat01":
//                        if (pVal.Row > 0)
//                        {
//                            oLastItemUID = pVal.ItemUID;
//                            oLastColUID = pVal.ColUID;
//                            oLastColRow = pVal.Row;
//                        }
//                        break;
//                    default:
//                        oLastItemUID = pVal.ItemUID;
//                        oLastColUID = "";
//                        oLastColRow = 0;
//                        break;
//                }
//            }
//            else if (pVal.BeforeAction == false)
//            {

//            }
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//            ////7
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//            ////8
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
//            ////9
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//            ////10
//            oForm.Freeze(true);
//            if (pVal.BeforeAction == true)
//            {

//            }
//            else if (pVal.BeforeAction == false)
//            {
//                if (pVal.ItemChanged == true)
//                {
//                    //                    If pVal.ItemUID = "Mat01" And pVal.ColUID = "" Then
//                    //                        Call PH_PY010_AddMatrixRow
//                    //                        Call oMat1.Columns(pVal.ColUID).Cells(pVal.Row).CLICK(ct_Regular)
//                    //                    End If
//                }
//            }
//            oForm.Freeze(false);
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//            ////11
//            if (pVal.BeforeAction == true)
//            {
//            }
//            else if (pVal.BeforeAction == false)
//            {
//                oMat1.LoadFromDataSource();

//                PH_PY010_FormItemEnabled();

//            }
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
//            ////12
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
//            ////16
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//            ////17
//            if (pVal.BeforeAction == true)
//            {
//            }
//            else if (pVal.BeforeAction == false)
//            {
//                SubMain.RemoveForms(oFormUniqueID);
//                //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                oForm = null;
//                //UPGRADE_NOTE: oDS_PH_PY010A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                oDS_PH_PY010A = null;


//                //UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                oMat1 = null;

//            }
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//            ////18
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//            ////19
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
//            ////20
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//            ////21
//            if (pVal.BeforeAction == true)
//            {

//            }
//            else if (pVal.BeforeAction == false)
//            {

//            }
//            break;
//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
//            ////22
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
//            ////23
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//            ////27
//            if (pVal.BeforeAction == true)
//            {

//            }
//            else if (pVal.Before_Action == false)
//            {
//                //                If pVal.ItemUID = "Code" Then
//                //                    Call MDC_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY010A", "Code")
//                //                End If
//            }
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
//            ////37
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
//            ////38
//            break;

//        //----------------------------------------------------------
//        case SAPbouiCOM.BoEventTypes.et_Drag:
//            ////39
//            break;

//    }

//    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oCombo = null;
//    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oRecordSet = null;

//    return;
//Raise_FormItemEvent_Error:
//    ///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//    oForm.Freeze((false));
//    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oCombo = null;
//    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    oRecordSet = null;
//    MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//}


//		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
//		{
//			int i = 0;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			oForm.Freeze(true);

//			if ((pVal.BeforeAction == true)) {
//				switch (pVal.MenuUID) {
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
//						break;
//				}
//			} else if ((pVal.BeforeAction == false)) {
//				switch (pVal.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY010_FormItemEnabled();
//						break;

//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY010_FormItemEnabled();

//						oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1282":
//						////문서추가
//						PH_PY010_FormItemEnabled();
//						break;

//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY010_FormItemEnabled();
//						break;
//					case "1293":
//						//// 행삭제
//						break;

//				}
//			}
//			oForm.Freeze(false);
//			return;
//			Raise_FormMenuEvent_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((BusinessObjectInfo.BeforeAction == true)) {
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
//			} else if ((BusinessObjectInfo.BeforeAction == false)) {
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
//			return;
//			Raise_FormDataEvent_Error:


//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//		}

//		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pVal.BeforeAction == true) {
//			} else if (pVal.BeforeAction == false) {
//			}
//			switch (pVal.ItemUID) {
//				case "Mat01":
//					if (pVal.Row > 0) {
//						oLastItemUID = pVal.ItemUID;
//						oLastColUID = pVal.ColUID;
//						oLastColRow = pVal.Row;
//					}
//					break;
//				default:
//					oLastItemUID = pVal.ItemUID;
//					oLastColUID = "";
//					oLastColRow = 0;
//					break;
//			}
//			return;
//			Raise_RightClickEvent_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}



//		public void PH_PY010_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY010'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//			}
//			return;
//			PH_PY010_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY010_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//private void PH_PY010_Proc()
//{
//    // ERROR: Not supported in C#: OnErrorStatement


//    string sQry = null;
//    string CLTCOD = null;
//    string DocDateFr = null;
//    string DocDateTo = null;
//    SAPbouiCOM.Form oForm = null;
//    SAPbobsCOM.Recordset oRecordSet = null;

//    SAPbouiCOM.ProgressBar ProgBar01 = null;

//    ProgBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("저장 중...", 100, false);

//    oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//    oForm = MDC_Globals.Sbo_Application.Forms.ActiveForm;

//    oForm.Freeze(true);

//    //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//    CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//    //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//    DocDateFr = oForm.Items.Item("DocDateFr").Specific.VALUE;
//    //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//    DocDateTo = oForm.Items.Item("DocDateTo").Specific.VALUE;

//    sQry = "PH_PY010_01 '" + CLTCOD + "','" + DocDateFr + "', '" + DocDateTo + "'";

//    oRecordSet.DoQuery(sQry);

//    oForm.Freeze(false);

//    ProgBar01.Value = 100;
//    ProgBar01.Stop();
//    //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    ProgBar01 = null;

//    MDC_Globals.Sbo_Application.MessageBox("근태처리 완료되었습니다.");
//    // StatusBar에 출력이 되지 않는 경우가 80% 이상 발생함, 완료메시지를 StatusBar가 아닌 MessageBox로 변경(2019.01.03 송명규)

//    return;
//Err_Renamed:


//    oForm.Freeze(false);

//    ProgBar01.Value = 100;
//    ProgBar01.Stop();
//    //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//    ProgBar01 = null;

//    MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY010_Proc_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//}
//	}
//}
