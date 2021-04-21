using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 여신한도 초과 승인
    /// </summary>
    internal class PS_SD082 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.DBDataSource oDS_PS_SD082L; //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_SD082M; //등록라인
        private string oLast_Item_UID; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLast_Col_UID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLast_Col_Row; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private int oLast_Mode;
        private int oSeq;

        /// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD082.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_SD082_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_SD082");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy = "DocNum";

                oForm.Freeze(true);

                PS_SD082_CreateItems();
                //PS_SD082_SetComboBox();
                //PS_SD082_Initialize();
                //oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                //PS_SD082_LoadCaption();

                oForm.EnableMenu("1283", false); //삭제
                oForm.EnableMenu("1286", false); //닫기
                oForm.EnableMenu("1287", false); //복제
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", false); //행삭제
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
        private void PS_SD082_CreateItems()
        {
            try
            {
                oDS_PS_SD082L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oDS_PS_SD082M = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat02 = oForm.Items.Item("Mat02").Specific;

                oForm.DataSources.UserDataSources.Add("Radio01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Radio01").Specific.DataBind.SetBound(true, "", "Radio01");

                oForm.DataSources.UserDataSources.Add("Radio02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Radio02").Specific.DataBind.SetBound(true, "", "Radio02");

                oForm.Items.Item("Radio01").Specific.GroupWith("Radio02");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }


        
        private void PS_SD082_SetComboBox()
        {

            
            SAPbouiCOM.ComboBox oCombo = null;
            string sQry = null;
            SAPbobsCOM.Recordset oRecordSet01 = null;

            oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oForm.DataSources.UserDataSources.Add("OkYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            oForm.Items.Item("OkYN").Specific.DataBind.SetBound(true, "", "OkYN");

            //// 승인상태
            //UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            oForm.Items.Item("OkYN").Specific.ValidValues.Add("Y", "승인");
            //UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            oForm.Items.Item("OkYN").Specific.ValidValues.Add("N", "미승인");
            //UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            oForm.Items.Item("OkYN").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_Index);

            //// 사업장
            oCombo = oForm.Items.Item("BPLId").Specific;
            sQry = "SELECT BPLId, BPLName From [OBPL] Order by BPLId";
            oRecordSet01.DoQuery(sQry);
            while (!(oRecordSet01.EoF))
            {
                oCombo.ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
                oMat01.Columns.Item("BPLId").ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
                oRecordSet01.MoveNext();
            }

            //// 사용자
            sQry = "Select empID, lastName + firstName From OHEM Order by empID";
            oRecordSet01.DoQuery(sQry);
            while (!(oRecordSet01.EoF))
            {
                oMat01.Columns.Item("CntcCode").ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
                oRecordSet01.MoveNext();
            }

        }







        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    int i = 0;
        //    int ErrNum = 0;
        //    object TempForm01 = null;
        //    SAPbouiCOM.ProgressBar ProgressBar01 = null;

        //    string ItemType = null;
        //    string RequestDate = null;
        //    string Size = null;
        //    string ItemCode = null;
        //    string ItemName = null;
        //    string Unit = null;
        //    string DueDate = null;
        //    string RequestNo = null;
        //    int Qty = 0;
        //    decimal Weight = default(decimal);
        //    string RFC_Sender = null;
        //    double Calculate_Weight = 0;
        //    int Seq = 0;

        //    ////BeforeAction = True
        //    if ((pVal.BeforeAction == true))
        //    {
        //        switch (pVal.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //                ////1
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //                ////2
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //                ////5
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_CLICK:
        //                ////6
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //                ////7
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //                ////8
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //                ////10
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //                ////11
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //                ////18
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //                ////19
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //                ////20
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //                ////27
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //                ////3
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //                ////4
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //                ////17
        //                break;
        //        }
        //        ////BeforeAction = False
        //    }
        //    else if ((pVal.BeforeAction == false))
        //    {
        //        switch (pVal.EventType)
        //        {
        //            //et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //                ////1
        //                if (pVal.ItemUID == "Btn01")
        //                {
        //                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //                    {
        //                        PS_SD082_UpdateSD080(ref pVal);
        //                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
        //                        oMat01.Clear();
        //                        oDS_PS_SD082L.Clear();
        //                        oMat02.Clear();
        //                        oDS_PS_SD082M.Clear();
        //                        PS_SD082_LoadCaption();
        //                    }
        //                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //                    {
        //                        oForm.Close();
        //                    }
        //                }
        //                else if (pVal.ItemUID == "Btn02")
        //                {
        //                    PS_SD082_LoadData();
        //                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
        //                    PS_SD082_LoadCaption();
        //                }
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //                ////2
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //                ////5
        //                break;
        //            //et_CLICK ///////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_CLICK:
        //                ////6
        //                if (pVal.ItemUID == "Radio01")
        //                {
        //                    oForm.Freeze(true);
        //                    oForm.Settings.MatrixUID = "Mat01";
        //                    oForm.Settings.EnableRowFormat = true;
        //                    oForm.Settings.Enabled = true;
        //                    oForm.Freeze(false);
        //                }
        //                else if (pVal.ItemUID == "Radio02")
        //                {
        //                    oForm.Freeze(true);
        //                    oForm.Settings.MatrixUID = "Mat02";
        //                    oForm.Settings.EnableRowFormat = true;
        //                    oForm.Settings.Enabled = true;
        //                    oForm.Freeze(false);
        //                }
        //                else if (pVal.ItemUID == "Mat01")
        //                {
        //                    if (pVal.ColUID == "LineNum")
        //                    {
        //                        PS_SD082_LoadData_Mat02((Strings.Trim(oDS_PS_SD082L.GetValue("U_ColReg02", pVal.Row - 1))));
        //                    }
        //                    else if (pVal.ColUID == "Check")
        //                    {
        //                        oForm.Freeze(true);
        //                        oMat01.FlushToDataSource();
        //                        for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
        //                        {
        //                            if (Strings.Trim(oDS_PS_SD082L.GetValue("U_ColReg01", i)) == "Y")
        //                            {
        //                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
        //                                PS_SD082_LoadCaption();
        //                                oForm.Freeze(false);
        //                                BubbleEvent = false;
        //                                return;
        //                            }
        //                        }
        //                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
        //                        PS_SD082_LoadCaption();
        //                        oForm.Freeze(false);
        //                        BubbleEvent = false;
        //                        return;
        //                    }
        //                }
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //                ////7
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //                ////8
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //                ////10
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //                ////11
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //                ////18
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //                ////19
        //                break;
        //            //et_FORM_RESIZE /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //                ////20
        //                oForm.Freeze(true);

        //                oForm.Items.Item("Mat01").Top = 50;
        //                oForm.Items.Item("Mat01").Left = 6;
        //                oForm.Items.Item("Mat01").Width = oForm.Width * 0.4 - 6;
        //                oForm.Items.Item("Mat01").Height = oForm.Height - 110;

        //                oForm.Items.Item("Mat02").Top = oForm.Items.Item("Mat01").Top;
        //                oForm.Items.Item("Mat02").Left = oForm.Width * 0.4 + 6 + 10;
        //                oForm.Items.Item("Mat02").Width = oForm.Width * 0.6 - 6 - 22;
        //                oForm.Items.Item("Mat02").Height = oForm.Height - 110;

        //                oForm.Items.Item("Radio01").Left = 6;
        //                oForm.Items.Item("Radio02").Left = oForm.Width * 0.4 + 6 + 10;

        //                oMat01.AutoResizeColumns();
        //                oMat02.AutoResizeColumns();

        //                //                oMat01.Columns("Check").Width = 40
        //                //                oMat01.Columns("DocNum").Width = 60
        //                //                oMat01.Columns("BPLId").Width = 50
        //                //                oMat01.Columns("CntcCode").Width = 60
        //                //                oMat01.Columns("DocDate").Width = 80
        //                //
        //                //                oMat02.Columns("CardCode").Width = 80
        //                //                oMat02.Columns("CardName").Width = 80
        //                //                oMat02.Columns("RequestP").Width = 80
        //                //                oMat02.Columns("CreditP").Width = 80
        //                //                oMat02.Columns("MiSuP").Width = 80
        //                //                oMat02.Columns("Balance").Width = 80
        //                //                oMat02.Columns("OutPreP").Width = 80
        //                //                oMat02.Columns("Comment").Width = 80

        //                oForm.Freeze(false);
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //                ////27
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //                ////3
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //                ////4
        //                break;
        //            //et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //                ////17
        //                SubMain.RemoveForms(oFormUniqueID01);
        //                //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oForm = null;
        //                //UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oMat01 = null;
        //                //UPGRADE_NOTE: oMat02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oMat02 = null;
        //                //UPGRADE_NOTE: oDS_PS_SD082L 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oDS_PS_SD082L = null;
        //                //UPGRADE_NOTE: oDS_PS_SD082M 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oDS_PS_SD082M = null;
        //                break;
        //        }
        //    }
        //    return;
        //Raise_ItemEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    //UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgressBar01 = null;
        //    if (ErrNum == 101)
        //    {
        //        ErrNum = 0;
        //        MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //        BubbleEvent = false;
        //    }
        //    else
        //    {
        //        MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //    }
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    int i = 0;

        //    ////BeforeAction = True
        //    if ((pVal.BeforeAction == true))
        //    {
        //        switch (pVal.MenuUID)
        //        {
        //            case "1284":
        //                //취소
        //                break;
        //            case "1286":
        //                //닫기
        //                break;
        //            case "1293":
        //                //행삭제
        //                break;
        //            case "1281":
        //                //찾기
        //                break;
        //            case "1282":
        //                //추가
        //                break;
        //            case "1288":
        //            case "1289":
        //            case "1290":
        //            case "1291":
        //                //레코드이동버튼
        //                break;
        //        }
        //        ////BeforeAction = False
        //    }
        //    else if ((pVal.BeforeAction == false))
        //    {
        //        switch (pVal.MenuUID)
        //        {
        //            case "1284":
        //                //취소
        //                break;
        //            case "1286":
        //                //닫기
        //                break;
        //            case "1293":
        //                //행삭제
        //                break;
        //            case "1281":
        //                //찾기
        //                break;
        //            case "1282":
        //                //추가
        //                break;
        //            case "1288":
        //            case "1289":
        //            case "1290":
        //            case "1291":
        //                //레코드이동버튼
        //                break;
        //        }
        //    }
        //    return;
        //Raise_MenuEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    MDC_Com.MDC_GF_Message(ref "Raise_MenuEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    ////BeforeAction = True
        //    if ((BusinessObjectInfo.BeforeAction == true))
        //    {
        //        switch (BusinessObjectInfo.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //                ////33
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //                ////34
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //                ////35
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //                ////36
        //                break;
        //        }
        //        ////BeforeAction = False
        //    }
        //    else if ((BusinessObjectInfo.BeforeAction == false))
        //    {
        //        switch (BusinessObjectInfo.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //                ////33
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //                ////34
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //                ////35
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //                ////36
        //                break;
        //        }
        //    }
        //    return;
        //Raise_FormDataEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    MDC_Com.MDC_GF_Message(ref "Raise_FormDataEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    if ((eventInfo.BeforeAction == true))
        //    {
        //        ////작업
        //    }
        //    else if ((eventInfo.BeforeAction == false))
        //    {
        //        ////작업
        //    }
        //    return;
        //Raise_RightClickEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion




        #region PS_SD082_Initialize
        //public void PS_SD082_Initialize()
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    SAPbouiCOM.ComboBox oCombo = null;

        //    ////아이디별 사업장 세팅
        //    oCombo = oForm.Items.Item("BPLId").Specific;
        //    oCombo.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_ByValue);

        //    ////아이디별 사번 세팅
        //    //    oForm.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD

        //    ////아이디별 부서 세팅
        //    //    Set oCombo = oForm.Items("DeptCode").Specific
        //    //    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
        //    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oCombo = null;
        //    return;
        //PS_SD082_Initialize_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oCombo = null;
        //    MDC_Com.MDC_GF_Message(ref "PS_SD082_Initialize_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region PS_SD082_LoadCaption
        //private void PS_SD082_LoadCaption()
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //    {
        //        //UPGRADE_WARNING: oForm.Items().Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        oForm.Items.Item("Btn01").Specific.Caption = "확인";
        //    }
        //    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //    {
        //        //UPGRADE_WARNING: oForm.Items().Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        oForm.Items.Item("Btn01").Specific.Caption = "확인";
        //    }
        //    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //    {
        //        //UPGRADE_WARNING: oForm.Items().Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        oForm.Items.Item("Btn01").Specific.Caption = "승인";
        //    }

        //    return;
        //PS_SD082_LoadCaption_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    MDC_Com.MDC_GF_Message(ref "Delete_EmptyRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region PS_SD082_LoadDataMat01
        //public void PS_SD082_LoadDataMat01()
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    short i = 0;
        //    string sQry = null;
        //    SAPbobsCOM.Recordset oRecordSet01 = null;
        //    oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    string OkYN = null;
        //    string BPLId = null;
        //    string DocNum = null;

        //    //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    BPLId = Strings.Trim(oForm.Items.Item("BPLId").Specific.Value);
        //    OkYN = Strings.Trim(oForm.DataSources.UserDataSources.Item("OkYN").Value);

        //    if (string.IsNullOrEmpty(OkYN))
        //        OkYN = "%";

        //    sQry = "EXEC [PS_SD082_01] '" + BPLId + "','" + OkYN + "','" + DocNum + "','01'";
        //    oRecordSet01.DoQuery(sQry);

        //    oMat01.Clear();
        //    oDS_PS_SD082L.Clear();

        //    oMat02.Clear();
        //    oDS_PS_SD082M.Clear();

        //    if (oRecordSet01.RecordCount == 0)
        //    {
        //        MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.:" + Err().Number + " - " + Err().Description, ref "W");
        //        //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //        oRecordSet01 = null;
        //        return;
        //    }

        //    oForm.Freeze(true);
        //    SAPbouiCOM.ProgressBar ProgBar01 = null;
        //    ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

        //    for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
        //    {
        //        if (i + 1 > oDS_PS_SD082L.Size)
        //        {
        //            oDS_PS_SD082L.InsertRecord((i));
        //        }

        //        oMat01.AddRow();
        //        oDS_PS_SD082L.Offset = i;
        //        oDS_PS_SD082L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
        //        oDS_PS_SD082L.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet01.Fields.Item("DocNum").Value));
        //        oDS_PS_SD082L.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet01.Fields.Item("U_BPLId").Value));
        //        oDS_PS_SD082L.SetValue("U_ColReg04", i, Strings.Trim(oRecordSet01.Fields.Item("U_CntcCode").Value));
        //        oDS_PS_SD082L.SetValue("U_ColReg05", i, Strings.Trim(oRecordSet01.Fields.Item("U_DocDate").Value));

        //        oRecordSet01.MoveNext();
        //        ProgBar01.Value = ProgBar01.Value + 1;
        //        ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
        //    }
        //    oMat01.LoadFromDataSource();
        //    oMat01.AutoResizeColumns();
        //    ProgBar01.Stop();
        //    oForm.Freeze(false);

        //    //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgBar01 = null;
        //    //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet01 = null;
        //    return;
        //PS_SD082_LoadDataMat01_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    ProgBar01.Stop();
        //    //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgBar01 = null;
        //    //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet01 = null;
        //    MDC_Com.MDC_GF_Message(ref "PS_SD082_LoadDataMat01_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region PS_SD082_LoadData_Mat02
        //public void PS_SD082_LoadDataMat02(string sDocNum)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    short i = 0;
        //    string sQry = null;
        //    SAPbobsCOM.Recordset oRecordSet01 = null;
        //    oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    string OkYN = null;
        //    string BPLId = null;
        //    string DocNum = null;

        //    //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    BPLId = Strings.Trim(oForm.Items.Item("BPLId").Specific.Value);
        //    OkYN = Strings.Trim(oForm.DataSources.UserDataSources.Item("OkYN").Value);

        //    sQry = "EXEC [PS_SD082_01] '" + BPLId + "','" + OkYN + "','" + sDocNum + "','02'";
        //    oRecordSet01.DoQuery(sQry);

        //    oMat02.Clear();
        //    oDS_PS_SD082M.Clear();

        //    if (oRecordSet01.RecordCount == 0)
        //    {
        //        MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.:" + Err().Number + " - " + Err().Description, ref "W");
        //        //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //        oRecordSet01 = null;
        //        return;
        //    }

        //    oForm.Freeze(true);
        //    SAPbouiCOM.ProgressBar ProgBar01 = null;
        //    ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

        //    for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
        //    {
        //        if (i + 1 > oDS_PS_SD082M.Size)
        //        {
        //            oDS_PS_SD082M.InsertRecord((i));
        //        }

        //        oMat02.AddRow();
        //        oDS_PS_SD082M.Offset = i;
        //        oDS_PS_SD082M.SetValue("U_LineNum", i, Convert.ToString(i + 1));
        //        oDS_PS_SD082M.SetValue("U_ColReg01", i, Strings.Trim(oRecordSet01.Fields.Item("U_CardCode").Value));
        //        oDS_PS_SD082M.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet01.Fields.Item("U_CardName").Value));
        //        oDS_PS_SD082M.SetValue("U_ColSum01", i, Strings.Trim(oRecordSet01.Fields.Item("U_RequestP").Value));
        //        oDS_PS_SD082M.SetValue("U_ColSum02", i, Strings.Trim(oRecordSet01.Fields.Item("U_CreditP").Value));
        //        oDS_PS_SD082M.SetValue("U_ColSum03", i, Strings.Trim(oRecordSet01.Fields.Item("U_MiSuP").Value));
        //        oDS_PS_SD082M.SetValue("U_ColSum04", i, Strings.Trim(oRecordSet01.Fields.Item("U_Balance").Value));
        //        oDS_PS_SD082M.SetValue("U_ColSum05", i, Strings.Trim(oRecordSet01.Fields.Item("U_OutPreP").Value));
        //        oDS_PS_SD082M.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet01.Fields.Item("U_Comment").Value));

        //        oRecordSet01.MoveNext();
        //        ProgBar01.Value = ProgBar01.Value + 1;
        //        ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
        //    }
        //    oMat02.LoadFromDataSource();
        //    oMat02.AutoResizeColumns();
        //    ProgBar01.Stop();
        //    oForm.Freeze(false);

        //    //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgBar01 = null;
        //    //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet01 = null;
        //    return;
        //PS_SD082_LoadDataMat02_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    ProgBar01.Stop();
        //    //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgBar01 = null;
        //    //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet01 = null;
        //    MDC_Com.MDC_GF_Message(ref "PS_SD082_LoadData_Mat02_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region PS_SD082_UpdateSD080
        //public bool PS_SD082_UpdateSD080(ref SAPbouiCOM.ItemEvent pVal)
        //{
        //    bool functionReturnValue = false;
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    short i = 0;
        //    string sQry = null;
        //    SAPbobsCOM.Recordset RecordSet01 = null;
        //    RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    string DocNum = null;
        //    string OkDate = null;

        //    oMat01.FlushToDataSource();

        //    for (i = 0; i <= oMat01.RowCount - 1; i++)
        //    {
        //        if (Strings.Trim(oDS_PS_SD082L.GetValue("U_ColReg01", i)) == "Y")
        //        {
        //            DocNum = Strings.Trim(oDS_PS_SD082L.GetValue("U_ColReg02", i));
        //            OkDate = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");

        //            sQry = "UPDATE [@PS_SD080H] ";
        //            sQry = sQry + "SET ";
        //            sQry = sQry + "U_OkYN = 'Y', ";
        //            sQry = sQry + "U_OkDate = '" + OkDate + "'";
        //            sQry = sQry + "Where DocNum = '" + DocNum + "'";

        //            RecordSet01.DoQuery(sQry);
        //        }
        //    }

        //    MDC_Com.MDC_GF_Message(ref "여신한도 초과승인 완료!", ref "S");

        //    //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet01 = null;
        //    return functionReturnValue;
        //Update_JakNum_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet01 = null;
        //    MDC_Com.MDC_GF_Message(ref "Update_JakNum_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //    return functionReturnValue;
        //}
        #endregion
    }
}
