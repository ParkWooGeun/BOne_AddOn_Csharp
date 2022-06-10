using System;
using SAPbouiCOM;

namespace PSH_BOne_AddOn.Core
{
    /// <summary>
    /// 계정과목표
    /// </summary>
    internal class S804 : PSH_BaseClass
    {
        private string oFormUniqueID;

        private string oLast_Item_UID; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLast_Col_UID;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLast_Col_Row;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="formUID"></param>
        public override void LoadForm(string formUID)
        {
            try
            {
                oFormUniqueID = formUID;
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);
                oForm.Freeze(true);
                S804_CreateItems();
                SubMain.Add_Forms(this, formUID, "S804");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// S804_CreateItems
        /// </summary>
        private void S804_CreateItems()
        {
            SAPbouiCOM.Item oNewITEM = null;
            int ComboWidth;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                ComboWidth = 50;

                oNewITEM = oForm.Items.Add("Text", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23;
                oNewITEM.Left = oForm.Items.Item("2005").Left;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = 60;
                oNewITEM.Specific.Caption = "분개전표";

                oNewITEM = oForm.Items.Add("RptCre01", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left + 39 + 51;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = ComboWidth;

                oNewITEM = oForm.Items.Add("RptDeb01", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left + 39;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = ComboWidth;

                oNewITEM = oForm.Items.Add("Text01", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = 40;
                oNewITEM.LinkTo = "RptCre01";
                oNewITEM.Specific.Caption = "항목1";

                oForm.Items.Item("RptCre01").Specific.DataBind.SetBound(true, "OACT", "U_RptCre01");
                oForm.Items.Item("RptDeb01").Specific.DataBind.SetBound(true, "OACT", "U_RptDeb01");

                oNewITEM = oForm.Items.Add("RptCre02", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left + 39 + 51;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = ComboWidth;

                oNewITEM = oForm.Items.Add("RptDeb02", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left + 39;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = ComboWidth;

                oNewITEM = oForm.Items.Add("Text02", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = 40;
                oNewITEM.LinkTo = "RptCre02";
                oNewITEM.Specific.Caption = "항목2";

                oForm.Items.Item("RptCre02").Specific.DataBind.SetBound(true, "OACT", "U_RptCre02");
                oForm.Items.Item("RptDeb02").Specific.DataBind.SetBound(true, "OACT", "U_RptDeb02");

                oNewITEM = oForm.Items.Add("RptCre03", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16 + 16 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left + 39 + 51;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = ComboWidth;

                oNewITEM = oForm.Items.Add("RptDeb03", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16 + 16 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left + 39;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = ComboWidth;

                oNewITEM = oForm.Items.Add("Text03", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16 + 16 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = 40;
                oNewITEM.LinkTo = "RptCre03";
                oNewITEM.Specific.Caption = "항목3";

                oForm.Items.Item("RptCre03").Specific.DataBind.SetBound(true, "OACT", "U_RptCre03");
                oForm.Items.Item("RptDeb03").Specific.DataBind.SetBound(true, "OACT", "U_RptDeb03");

                oNewITEM = oForm.Items.Add("RptCre04", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left + 39 + 51 + 4 + 40 + 50 + 50 + 12;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = ComboWidth;

                oNewITEM = oForm.Items.Add("RptDeb04", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left + 39 + 51 + 4 + 40 + 50 + 12;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = ComboWidth;

                oNewITEM = oForm.Items.Add("Text04", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left + 39 + 51 + 4 + 50 + 12;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = 40;
                oNewITEM.LinkTo = "RptCre04";
                oNewITEM.Specific.Caption = "항목4";

                oForm.Items.Item("RptCre04").Specific.DataBind.SetBound(true, "OACT", "U_RptCre04");
                oForm.Items.Item("RptDeb04").Specific.DataBind.SetBound(true, "OACT", "U_RptDeb04");

                oNewITEM = oForm.Items.Add("RptCre05", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left + 39 + 51 + 4 + 40 + 50 + 50 + 12;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = ComboWidth;

                oNewITEM = oForm.Items.Add("RptDeb05", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left + 39 + 51 + 4 + 40 + 50 + 12;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = ComboWidth;

                oNewITEM = oForm.Items.Add("Text05", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left + 39 + 51 + 4 + 50 + 12;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = 40;
                oNewITEM.LinkTo = "RptCre05";
                oNewITEM.Specific.Caption = "항목5";

                oForm.Items.Item("RptCre05").Specific.DataBind.SetBound(true, "OACT", "U_RptCre05");
                oForm.Items.Item("RptDeb05").Specific.DataBind.SetBound(true, "OACT", "U_RptDeb05");

                oNewITEM = oForm.Items.Add("RptCre06", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16 + 16 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left + 39 + 51 + 4 + 40 + 50 + 50 + 12;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = ComboWidth;

                oNewITEM = oForm.Items.Add("RptDeb06", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16 + 16 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left + 39 + 51 + 4 + 40 + 50 + 12;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = ComboWidth;

                oNewITEM = oForm.Items.Add("Text06", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Top = oForm.Items.Item("2005").Top + 23 + 16 + 16 + 16;
                oNewITEM.Left = oForm.Items.Item("2005").Left + 39 + 51 + 4 + 50 + 12;
                oNewITEM.Height = oForm.Items.Item("2005").Height;
                oNewITEM.Width = 40;
                oNewITEM.LinkTo = "RptCre06";
                oNewITEM.Specific.Caption = "항목6";

                oForm.Items.Item("RptCre06").Specific.DataBind.SetBound(true, "OACT", "U_RptCre06");
                oForm.Items.Item("RptDeb06").Specific.DataBind.SetBound(true, "OACT", "U_RptDeb06");

                //Combo
                sQry = "select U_Minor, U_CdName from [@PS_SY001L] Where Code = 'F001' Order by Convert(Int, U_LineNum)";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("RptCre01").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oForm.Items.Item("RptDeb01").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oForm.Items.Item("RptCre02").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oForm.Items.Item("RptDeb02").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oForm.Items.Item("RptCre03").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oForm.Items.Item("RptDeb03").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oForm.Items.Item("RptCre04").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oForm.Items.Item("RptDeb04").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oForm.Items.Item("RptCre05").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oForm.Items.Item("RptDeb05").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oForm.Items.Item("RptCre06").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oForm.Items.Item("RptDeb06").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }

                oNewITEM = oForm.Items.Add("AddonText", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Top = oForm.Items.Item("4").Top - 12;
                oNewITEM.Left = oForm.Items.Item("4").Left;
                oNewITEM.Height = 12;
                oNewITEM.Width = 120;
                oNewITEM.FontSize = 10;
                oNewITEM.Specific.Caption = "Addon running";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oNewITEM);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                //case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                //	Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //	Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //	Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //	Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //	break;
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
                //	Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //	Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //	Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //	break;
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
                //	Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //	break;
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
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
            }
        }

        /// <summary>
        /// Raise_EVENT_GOT_FOCUS
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    oLast_Item_UID = pVal.ItemUID;

                    if (oLast_Item_UID == "38")
                    {
                        if (pVal.Row > 0)
                        {
                            oLast_Item_UID = pVal.ItemUID;
                            oLast_Col_UID = pVal.ColUID;
                            oLast_Col_Row = pVal.Row;
                        }
                    }
                    else
                    {
                        oLast_Item_UID = pVal.ItemUID;
                        oLast_Col_UID = "";
                        oLast_Col_Row = 0;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    oLast_Item_UID = pVal.ItemUID;

                    if (oLast_Item_UID == "38")
                    {
                        if (pVal.Row > 0)
                        {
                            oLast_Item_UID = pVal.ItemUID;
                            oLast_Col_UID = pVal.ColUID;
                            oLast_Col_Row = pVal.Row;
                        }
                    }
                    else
                    {
                        oLast_Item_UID = pVal.ItemUID;
                        oLast_Col_UID = "";
                        oLast_Col_Row = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1283": //삭제
                            break;
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "1293": //행삭제
                            break;
                        case "7169": //엑셀 내보내기
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1287": // 복제
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "1293": //행삭제
                            break;
                        case "7169": //엑셀 내보내기
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
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }
    }
}
