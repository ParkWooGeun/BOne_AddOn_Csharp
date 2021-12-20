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
        private string itemCode;
        //private string frDate;
        //private string toDate;
        private string chkValue;
        private BoFormMode formMode; //Form.Mode 저장

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
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void S150_CreateItems()
        {
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Item oItem_ItmMsort = null;
            SAPbouiCOM.Item oItem_Spec2 = null;
            SAPbouiCOM.Item oItem_Spec4 = null;
            SAPbouiCOM.Item oCombo = null;

            try
            {
                oItem = oForm.Items.Add("Text", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Top = oForm.Items.Item("10002052").Top + 23;
                oItem.Left = oForm.Items.Item("10002052").Left + 20;
                oItem.Height = oForm.Items.Item("10002052").Height;
                oItem.Width = 80;
                oItem.Specific.Caption = "코드사용여부";
                oItem.FromPane = 6;
                oItem.ToPane = 6;

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
                oCombo.FromPane = 6;
                oCombo.ToPane = 6;

                oItem = oForm.Items.Add("AddonText", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Top = oForm.Items.Item("1").Top - 12;
                oItem.Left = oForm.Items.Item("1").Left;
                oItem.Height = 12;
                oItem.Width = 120;
                oItem.FontSize = 10;
                oItem.Specific.Caption = "Addon running";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem_ItmMsort);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem_Spec2);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem_Spec4);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCombo);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void S150_FormItemEnabled()
        {
            try
            {
                //Form Mode에 상관없이 무조건 아래 컨트롤 비활성
                oForm.Items.Item("10002050").Enabled = false; //활성(라디오)
                oForm.Items.Item("10002051").Enabled = false; //비활성(라디오)
                oForm.Items.Item("10002052").Enabled = false; //고급(라디오)

                oForm.Items.Item("10002045").Enabled = false; //시작(비활성)(일자)
                oForm.Items.Item("10002042").Enabled = false; //종료(비활성)(일자)
                oForm.Items.Item("10002047").Enabled = false; //비고(비활성)(텍스트)

                oForm.Items.Item("10002038").Enabled = false; //시작(활성)(일자)
                oForm.Items.Item("10002041").Enabled = false; //종료(활성)(일자)
                oForm.Items.Item("10002048").Enabled = false; //비고(활성)(텍스트)
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
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool S150_DataValidCheck()
        {       
            bool returnValue = false;
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
                        errMessage += "뷰 > 사용자정의필드를 선택하여 규격2,4 필드를 숫자로 변경하세요.";
                        throw new Exception();
                    }
                }

                if (oForm.Items.Item("10002047").Specific.Value == "미사용") //비활성 라디오버튼의 비고가 "미사용"이면
                {
                    if (oForm.Items.Item("CheckYN").Specific.Value != "Y") //사용 선택은 가능
                    {
                        errMessage = "미사용 상태에서는 수정할 수 없습니다.";
                        throw new Exception();
                    }
                }
                else if (oForm.Items.Item("10002047").Specific.Value.ToString().Trim() != "" && oForm.Items.Item("10002047").Specific.Value.ToString().Trim().Split('-')[1] == "승인필요") //승인전이면
                {
                    errMessage = "승인전 상태에서는 수정할 수 없습니다.";
                    throw new Exception();
                }

                returnValue = true;
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
            return returnValue;
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool S150_Power_Check()
        {
            bool returnValue = false;
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

                returnValue = true;
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
            return returnValue;
        }

        /// <summary>
        /// 추가/수정 시 승인을 위한 상태 변경
        /// </summary>
        private void S150_SetAuthorityStatus()
        {
            string sQry;
            string message = string.Empty;
            string frDate;
            string toDate = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                frDate = DateTime.Now.ToString("yyyyMMdd");

                if (formMode == BoFormMode.fm_UPDATE_MODE || formMode == BoFormMode.fm_ADD_MODE) //추가/수정 모드에서만 동작
                {
                    if (formMode == BoFormMode.fm_ADD_MODE) //추가(최초등록)
                    {
                        toDate = "29991231";
                        message = "최초등록-승인필요";
                    }
                    else if (formMode == BoFormMode.fm_UPDATE_MODE) //수정(수정등록, 사용등록)
                    {
                        if (chkValue == "") //사용여부-미선택
                        {
                            if (oForm.Items.Item("10002050").Specific.Selected == true) //활성 라디오 버튼 선택된 경우 (승인후)
                            {
                                toDate = "29991231";
                                message = "수정등록-승인필요";
                                oForm.Items.Item("10002051").Enabled = true; //비활성(라디오)
                                oForm.Items.Item("10002051").Specific.Selected = true;
                                oForm.Items.Item("10002051").Enabled = false; //비활성(라디오)
                            }
                        }
                        else if (chkValue == "Y") //사용여부-사용
                        {
                            toDate = "29991231";
                            message = "사용등록-승인필요";
                            oForm.Items.Item("10002051").Enabled = true; //비활성(라디오)
                            oForm.Items.Item("10002051").Specific.Selected = true;
                            oForm.Items.Item("10002051").Enabled = false; //비활성(라디오)
                        }
                        else if (chkValue == "N") //사용여부-미사용
                        {
                            toDate = "28991231";
                            message = "미사용";
                            oForm.Items.Item("10002051").Enabled = true; //비활성(라디오)
                            oForm.Items.Item("10002051").Specific.Selected = true;
                            oForm.Items.Item("10002051").Enabled = false; //비활성(라디오)
                        }
                    }

                    sQry = "  UPDATE    OITM";
                    sQry += " SET       UpdateDate = '" + frDate + "',";
                    sQry += "           validFor = 'N',";
                    sQry += "           frozenFor = 'Y',";
                    sQry += "           frozenFrom = '" + frDate + "',";
                    sQry += "           frozenTo = '" + toDate + "',";
                    sQry += "           FrozenComm = '" + message + "'";
                    sQry += " FROM      OITM";
                    sQry += " WHERE     ItemCode = '" + itemCode + "'";

                    oRecordSet.DoQuery(sQry);
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
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "12" || pVal.ItemUID == "13")
                    {
                        sQry = "select U_Module from [@PS_SY005L]  where Code ='S150' and U_UseYN ='Y' and U_USERID ='" + PSH_Globals.oCompany.UserName + "'";
                        oRecordSet01.DoQuery(sQry);

                        if (oRecordSet01.Fields.Item(0).Value == "M2" || string.IsNullOrEmpty(oRecordSet01.Fields.Item(0).Value))
                        {
                            errCode = 1;
                            throw new Exception();
                        }
                        else if (oRecordSet01.Fields.Item(0).Value == "M3" || string.IsNullOrEmpty(oRecordSet01.Fields.Item(0).Value))
                        {
                            errCode = 2;
                            throw new Exception();
                        }
                    }
                    else if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
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

                            //추가,수정 후 BeforeAction == false로 전달할 데이터_S
                            itemCode = oForm.Items.Item("5").Specific.Value;
                            chkValue = oForm.Items.Item("CheckYN").Specific.Value.ToString().Trim();
                            //추가,수정 후 BeforeAction == false로 전달할 데이터_E
                            oForm.Items.Item("CheckYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }

                        formMode = oForm.Mode; //BeforeAction == false로 전달할 Form Mode
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {   
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            S150_FormItemEnabled();
                        }

                        if (pVal.ActionSuccess == true)
                        {
                            S150_SetAuthorityStatus();
                        }

                        oForm.Items.Item("7").Click(); //품목명 클릭
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
            string errMessage = string.Empty;

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
                        if (oForm.Items.Item("10002047").Specific.Value != "미사용")
                        {
                            errMessage = "선택 불가";
                            oForm.Items.Item("CheckYN").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index); //기본값
                            throw new Exception();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                S150_FormItemEnabled();
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {   
                }
                else if (pVal.Before_Action == false)
                {
                    S150_FormItemEnabled();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
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
                            S150_FormItemEnabled();
                            break;
                        case "1282": //추가
                            S150_FormItemEnabled();
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
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
