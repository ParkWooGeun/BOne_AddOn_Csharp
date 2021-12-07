using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Code;
using System.Collections.Generic;

namespace PSH_BOne_AddOn.Core
{

    /// <summary>
    /// 일반 권한
    /// </summary>
    internal class S951 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private string cUserID;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="formUID"></param>
        public override void LoadForm(string formUID)
        {
            try
            {
                oForm = PSH_Globals.SBO_Application.Forms.Item(formUID);
                oForm.Freeze(true);

                oFormUniqueID = formUID;
                SubMain.Add_Forms(this, formUID, "S951");

                S951_CreateItems();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void S951_CreateItems()
        {
            SAPbouiCOM.Item oItem;

            try
            {
                oMat01 = oForm.Items.Item("6").Specific;
                oMat02 = oForm.Items.Item("5").Specific;

                oItem = oForm.Items.Add("AddonText", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Top = oForm.Items.Item("1").Top - 12;
                oItem.Left = oForm.Items.Item("1").Left;
                oItem.Height = 12;
                oItem.Width = 120;
                oItem.FontSize = 10;
                oItem.Specific.Caption = "Addon running";

                oItem = oForm.Items.Add("refText", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Top = oForm.Items.Item("2").Top;
                oItem.Left = oForm.Items.Item("2").Left + 80;
                oItem.Height = 14;
                oItem.Width = 80;
                oItem.LinkTo = "refEdit";
                oItem.Specific.Caption = "관련근거";

                oItem = oForm.Items.Add("refEdit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Top = oForm.Items.Item("2").Top;
                oItem.Left = oForm.Items.Item("2").Left + 160;
                oItem.Height = 14;
                oItem.Width = 180;

                oItem = oForm.Items.Add("MtypeL", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Top = oForm.Items.Item("2").Top;
                oItem.Left = oForm.Items.Item("refEdit").Left + 190;
                oItem.Height = 14;
                oItem.Width = 80;
                oItem.LinkTo = "MType";
                oItem.Specific.Caption = "변경타입";

                oItem = oForm.Items.Add("MType", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oItem.Top = oForm.Items.Item("2").Top;
                oItem.Left = oForm.Items.Item("MtypeL").Left + 80;
                oItem.Specific.ValidValues.Add("N", "신규");
                oItem.Specific.ValidValues.Add("M", "변경");
                oItem.Specific.ValidValues.Add("C", "부서이동");
                oItem.Height = 14;
                oItem.Width = 180;
                oItem.DisplayDesc = true;
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
        /// SaveLogData
        /// </summary>
        private bool S951_SaveLogData()
        {
            bool returnValue = false;
            int i;
            string sQry;
            S230 s230 = new S230();
            string addString = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("저장중", 0, false);
                if (s230.RegUserID.Count > 0) // 복제
                {
                    oForm.Items.Item("13").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    for (i = 1; i <= oMat01.VisualRowCount; i++)
                    {
                        addString += oMat01.Columns.Item("1").Cells.Item(i).Specific.Value + "^" + oMat01.Columns.Item("2").Cells.Item(i).Specific.Value + "`";
                    }
                    sQry = "select top 1 data   from ps_sy020 where UserID ='"+ cUserID + "'order by CreateDate desc";
                    oRecordSet01.DoQuery(sQry);
                    if (oRecordSet01.Fields.Item(0).Value != addString) // 복사 대상에 수정사항이 발생하면 복사대상도 List에 추가시켜서 Log 저장
                    {
                        s230.RegUserID.Add(cUserID);
                    }
                    for (i = 0; i < s230.RegUserID.Count; i++)
                    {
                        sQry = "Insert into PS_SY020 SELECT 'C','" + s230.RegUserID[i] + "','" + addString + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + PSH_Globals.oCompany.UserName.ToString() + "','";
                        sQry += oForm.Items.Item("refEdit").Specific.Value.ToString().Trim() + "','N','"+ oForm.Items.Item("MType").Specific.Value +"'";
                        oRecordSet01.DoQuery(sQry);
                        if (oForm.Items.Item("refEdit").Specific.Value.ToString().Trim() != "초기등록")
                        {
                            sQry = "Exec PS_SY020_02 '" + s230.RegUserID[i] + "'";
                            oRecordSet01.DoQuery(sQry);
                        }
                        else
                        {
                            sQry = "UPDATE PS_SY020 set Status = 'Y' WHERE UserID ='" + s230.RegUserID[i] + "'";
                            oRecordSet01.DoQuery(sQry);
                        }
                    }
                    s230.RegUserID.Clear();
                }
                else 
                {
                    oForm.Items.Item("13").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    for (i = 1; i <= oMat01.VisualRowCount; i++)
                    {
                        addString += oMat01.Columns.Item("1").Cells.Item(i).Specific.Value + "^" + oMat01.Columns.Item("2").Cells.Item(i).Specific.Value + "`";//1
                    }
                    sQry = "Insert into PS_SY020 SELECT 'C','" + cUserID + "','" + addString + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + PSH_Globals.oCompany.UserName.ToString() + "','";
                    sQry += oForm.Items.Item("refEdit").Specific.Value.ToString().Trim() + "','N','" + oForm.Items.Item("MType").Specific.Value + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (oForm.Items.Item("refEdit").Specific.Value.ToString().Trim() != "초기등록")
                    {
                        sQry = "Exec PS_SY020_02 '" + cUserID + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                    else
                    {
                        sQry = "UPDATE PS_SY020 set Status = 'Y' WHERE UserID ='" + cUserID + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                }
            returnValue = true;
            }
            catch (Exception ex)
            {
            }
            finally
            {
                if (ProgBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
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
                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                    //Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("refEdit").Specific.Value.ToString()))
                        {
                            errMessage = "관련근거는 필수 입력입니다";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oForm.Items.Item("MType").Specific.Value.ToString()))
                        {
                            errMessage = "변경타입은 필수 입력입니다.";
                            throw new Exception();
                        }
                        else
                        {
                            if (S951_SaveLogData() == false)
                            {
                                errMessage = "수정중 오류발생";
                                throw new Exception();
                            }
                            else
                            {
                                PSH_Globals.SBO_Application.MessageBox("수정완료");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if(errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
                BubbleEvent = false;
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
                    if(pVal.ItemUID == "5")
                    {
                        cUserID = oMat02.Columns.Item("0").Cells.Item(pVal.Row).Specific.Value;
                    }
                }
                else if (pVal.Before_Action == false)
                {
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
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
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
                            break;
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                            break;
                        case "1287": //복제
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
            finally
            {
            }
        }
    }
}
