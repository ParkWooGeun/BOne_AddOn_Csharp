using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn.Core
{
    /// <summary>
    /// 품목마스터
    /// </summary>
    internal class S150 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private string ItemCode;
        private string FrDate;
        private string ToDate;
        private string BeMode;//추가/갱신인지 확인

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="formUID"></param>
        public override void LoadForm(string formUID)
        {
            try
            {
                oForm = PSH_Globals.SBO_Application.Forms.Item(formUID);
                oFormUniqueID = formUID;
                SubMain.Add_Forms(this, formUID, "S150");
                
                S150_CreateItems();
                S150_FormItemEnabled();
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
        private void S150_CreateItems()
        {
            SAPbouiCOM.Item oItem;
            SAPbouiCOM.Item oItem_ItmMsort;
            SAPbouiCOM.Item oItem_Spec2;
            SAPbouiCOM.Item oItem_Spec4;
            SAPbouiCOM.Item oCombo;

            try
            {
                oForm.Freeze(true);
                oItem = oForm.Items.Add("Text", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Top = oForm.Items.Item("10002052").Top + 23;
                oItem.Left = oForm.Items.Item("10002052").Left + 20;
                oItem.Height = oForm.Items.Item("10002052").Height;
                oItem.Width = 80;
                oItem.Specific.Caption = "코드사용여부";

                oItem_ItmMsort = oForm.Items.Add("ItmMsort", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem_ItmMsort.Left = 250;
                oItem_ItmMsort.Top = oItem.Top + 23;
                oItem_ItmMsort.Height = oItem.Height;
                oItem_ItmMsort.Width = 80;
                oItem_ItmMsort.Visible = false;
                SAPbouiCOM.EditText oEdit_ItmMsort = oItem_ItmMsort.Specific;
                oEdit_ItmMsort.DataBind.SetBound(true, "OITM", "U_ItmMsort");

                oItem_Spec2 = oForm.Items.Add("Spec2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem_Spec2.Left = 250;
                oItem_Spec2.Top = oItem.Top + 23;
                oItem_Spec2.Height = oItem.Height;
                oItem_Spec2.Width = 80;
                oItem_Spec2.Visible = false;
                SAPbouiCOM.EditText oEdit_Spec2 = oItem_Spec2.Specific;
                oEdit_Spec2.DataBind.SetBound(true, "OITM", "U_Spec2");

                oItem_Spec4 = oForm.Items.Add("Spec4", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem_Spec4.Left = 250;
                oItem_Spec4.Top = oItem.Top + 23;
                oItem_Spec4.Height = oItem.Height;
                oItem_Spec4.Width = 80;
                oItem_Spec4.Visible = false;
                SAPbouiCOM.EditText oEdit_Spec4 = oItem_Spec4.Specific;
                oEdit_Spec4.DataBind.SetBound(true, "OITM", "U_Spec4");

                oCombo = oForm.Items.Add("CheckYN", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oCombo.Left = 120;
                oCombo.Top = oForm.Items.Item("10002052").Top + 23;
                oCombo.Height = oForm.Items.Item("10002052").Height;
                oCombo.Width = 80;
                oCombo.DisplayDesc = true;
                oCombo.Specific.ValidValues.Add("", "-");
                oCombo.Specific.ValidValues.Add("Y", "사용");
                oCombo.Specific.ValidValues.Add("N", "미사용");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void S150_FormItemEnabled()
        {
            try
            {
                oForm.Items.Item("10002050").Enabled = false; //활성
                oForm.Items.Item("10002051").Enabled = false; //비활성
                oForm.Items.Item("10002052").Enabled = false; //고급
                oForm.Items.Item("10002045").Enabled = false; //시작(비)
                oForm.Items.Item("10002042").Enabled = false; //종료(비)
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
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool S150_DataValidCheck()
        {       
            bool functionReturnValue = false;
            double chknum;
            string errMessage = string.Empty;

            try
            {
                if (oForm.Items.Item("39").Specific.Value == "105")
                {
                    errMessage = "저장품(부자재)은 신규등록/갱신 할 수 없습니다.";
                    throw new Exception();
                }
                if (oForm.Items.Item("ItmMsort").Specific.Value == "30603") // 중분류가 봉일 경우 Spec2, Spec4 필드에 문자, 공백 불가 로직 추가
                {
                    if((double.TryParse(oForm.Items.Item("Spec2").Specific.Value, out chknum) == false || double.TryParse(oForm.Items.Item("Spec4").Specific.Value, out chknum) == false) || (string.IsNullOrEmpty(oForm.Items.Item("Spec2").Specific.Value.ToString().Trim()) || string.IsNullOrEmpty(oForm.Items.Item("Spec4").Specific.Value.ToString().Trim())))
                    {
                        errMessage = "오류 : 중분류 봉(30603)의 경우 규격2,4 필드에 문자, 공백 불가!\n";
                        errMessage += "뷰 > 사용자정의필드를 선택하여 규격2,4필드를 숫자로 변경하세요.";
                        throw new Exception();
                    }
                }
                functionReturnValue = true;
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
            }
            return functionReturnValue;
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool S150_Power_Check()
        {
            bool functionReturnValue = false;
            string sQry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "select U_Power from [@PS_SY005L] where 1=1 and code ='S150' and U_UseYN ='Y' and U_USERID ='" + PSH_Globals.oCompany.UserName + "'";
                oRecordSet01.DoQuery(sQry);

                if (string.IsNullOrEmpty(oRecordSet01.Fields.Item(0).Value))
                {
                    errMessage = "읽기 사용자 : 추가 수정 불가";
                    throw new Exception();
                }

                functionReturnValue = true;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return functionReturnValue;
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

                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
            string sQry;
            int errCode = 0;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "12")
                    {
                        sQry = "select U_Module from [@PS_SY005L]  where Code ='S150' and U_UseYN ='Y' and U_USERID ='" + PSH_Globals.oCompany.UserName + "'";
                        oRecordSet01.DoQuery((sQry));

                        if (oRecordSet01.Fields.Item(0).Value == "M2" || string.IsNullOrEmpty(oRecordSet01.Fields.Item(0).Value))
                        {
                            errCode = 1;
                            throw new Exception();
                        }
                    }
                    else if (pVal.ItemUID == "13")
                    {
                        sQry = "select U_Module from [@PS_SY005L]  where Code ='S150' and U_UseYN ='Y' and U_USERID ='" + PSH_Globals.oCompany.UserName + "'";
                        oRecordSet01.DoQuery((sQry));

                        if (oRecordSet01.Fields.Item(0).Value == "M3" || string.IsNullOrEmpty(oRecordSet01.Fields.Item(0).Value))
                        {
                            errCode = 2;
                            throw new Exception();
                        }
                    }

                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (S150_Power_Check() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (S150_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            FrDate = DateTime.Now.ToString("yyyyMMdd");

                            if (oForm.Items.Item("CheckYN").Specific.Value.ToString().Trim() == "N")
                            {
                                ToDate = "28991231";
                            }
                            else
                            {
                                ToDate = "29991231";
                            }
                            ItemCode = oForm.Items.Item("5").Specific.Value;

                            BeMode = Convert.ToString(oForm.Mode);
                            oForm.Items.Item("CheckYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (S150_Power_Check() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (S150_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            FrDate = DateTime.Now.ToString("yyyyMMdd");

                            if (oForm.Items.Item("CheckYN").Specific.Value.ToString().Trim() == "N")
                            {
                                ToDate = "28991231";
                            }
                            else
                            {
                                ToDate = "29991231";
                            }
                            ItemCode = oForm.Items.Item("5").Specific.Value;

                            BeMode = Convert.ToString(oForm.Mode);
                            oForm.Items.Item("CheckYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
                                S150_FormItemEnabled();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                S150_FormItemEnabled();
                            }
                        }
                        if (pVal.ActionSuccess == true)
                        {
                            if (oForm.Items.Item("10002051").Specific.Selected == false)
                            {
                                if (oForm.Items.Item("10002047").Specific.Value != "초기등록")
                                {
                                    if (BeMode == "2") // 업데이트시 마지막에 업데이트 일자 처리
                                    {
                                        sQry = "update oitm set UpdateDate='" + FrDate + "', validFor = 'N', frozenFor    = 'Y', frozenFrom = '" + FrDate + "', frozenTo = '" + ToDate + "', FrozenComm = '업데이트됨' FROM OITM WHERE ITEMCODE ='" + ItemCode + "'";
                                        oRecordSet01.DoQuery(sQry);
                                        BeMode = "0";
                                    }
                                    else if (BeMode == "3") // 신규추가시 마지막에 업데이트 일자 처리
                                    {
                                        sQry = "update oitm set UpdateDate='" + FrDate + "', validFor = 'N', frozenFor    = 'Y', frozenFrom = '" + FrDate + "', frozenTo = '" + ToDate + "', FrozenComm = '초기등록' FROM OITM WHERE ITEMCODE ='" + ItemCode + "'";
                                        oRecordSet01.DoQuery(sQry);
                                        BeMode = "0";
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (errCode == 1)
                {
                    if (oForm.Items.Item("12").Specific.Checked == true)
                    {
                        PSH_Globals.SBO_Application.MessageBox("해당 권한으로 구매 품목은 선택불가합니다.");
                        oForm.Items.Item("12").Specific.Checked = false;
                    }
                }
                else if (errCode == 2)
                {
                    if (oForm.Items.Item("13").Specific.Checked == true)
                    {
                        PSH_Globals.SBO_Application.MessageBox("해당 권한으로 판매 품목은 선택불가합니다.");
                        oForm.Items.Item("13").Specific.Checked = false;
                    }
                }
                else if (errCode == 3)
                {
                    PSH_Globals.SBO_Application.MessageBox("해당 권한으로 선택불가합니다.");
                    sQry = "select ItmsGrpCod  from [OITM] where itemcode ='" + oForm.Items.Item("5").Specific.Value + "'";
                    oRecordSet01.DoQuery((sQry));
                    oForm.Items.Item("39").Specific.Select(Convert.ToString(Convert.ToDouble(codeHelpClass.Right(oRecordSet01.Fields.Item(0).Value, 1)) - 1), SAPbouiCOM.BoSearchKey.psk_Index);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
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
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (oForm.Items.Item("CheckYN").Specific.Value.ToString().Trim() == "Y")
                    {
                        oForm.Items.Item("10002050").Enabled = true; //활성
                        oForm.Items.Item("10002050").Specific.Selected = true;
                        oForm.Items.Item("10002050").Enabled = false; //활성
                        oForm.Items.Item("10002051").Enabled = false; //비활성
                        oForm.Items.Item("10002052").Enabled = false; //고급
                    }
                    else if (oForm.Items.Item("CheckYN").Specific.Value.ToString().Trim() == "N")
                    {
                        oForm.Items.Item("10002051").Enabled = true; //비활성
                        oForm.Items.Item("10002051").Specific.Selected = true;
                        oForm.Items.Item("10002050").Enabled = false; //활성
                        oForm.Items.Item("10002051").Enabled = false; //비활성
                        oForm.Items.Item("10002052").Enabled = false;  //고급
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int errCode = 0;
            string sQry;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Items.Item("10002050").Enabled = false; //활성
                oForm.Items.Item("10002051").Enabled = false; //비활성
                oForm.Items.Item("10002052").Enabled = false; //고급
                oForm.Items.Item("10002047").Enabled = false; //고급

                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "39")
                    {
                        sQry = "select U_Module from [@PS_SY005L]  where Code ='S150' and U_UseYN ='Y' and U_USERID ='" + PSH_Globals.oCompany.UserName + "'";
                        oRecordSet01.DoQuery(sQry);

                        if (oRecordSet01.Fields.Item(0).Value == "M3")
                        {
                            if ((oForm.Items.Item("39").Specific.Selected.Value.ToString().Trim() == "102") 
                             || (oForm.Items.Item("39").Specific.Selected.Value.ToString().Trim() == "106") 
                             || (oForm.Items.Item("39").Specific.Selected.Value.ToString().Trim() == "103"))
                            {
                                errCode = 3;
                                throw new Exception();
                            }
                        }
                        else if (oRecordSet01.Fields.Item(0).Value == "M2")
                        {
                            if ((oForm.Items.Item("39").Specific.Selected.Value.ToString().Trim() == "101") 
                             || (oForm.Items.Item("39").Specific.Selected.Value.ToString().Trim() == "104") 
                             || (oForm.Items.Item("39").Specific.Selected.Value.ToString().Trim() == "105"))
                            {
                                errCode = 3;
                                throw new Exception();
                            }
                        }
                        else if (string.IsNullOrEmpty(oRecordSet01.Fields.Item(0).Value))
                        {
                            errCode = 3;
                            throw new Exception();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (errCode == 3)
                {
                    PSH_Globals.SBO_Application.MessageBox("해당 권한으로 선택불가합니다.");
                    sQry = "select ItmsGrpCod  from [OITM] where itemcode ='" + oForm.Items.Item("5").Specific.Value + "'";
                    oRecordSet01.DoQuery(sQry);
                    oForm.Items.Item("39").Specific.Select(Convert.ToString(Convert.ToDouble(codeHelpClass.Right(oRecordSet01.Fields.Item(0).Value, 1)) - 1), SAPbouiCOM.BoSearchKey.psk_Index);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
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
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
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
                            S150_FormItemEnabled();
                            break;
                        case "1282": //추가
                            S150_FormItemEnabled();
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                            S150_FormItemEnabled();
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
