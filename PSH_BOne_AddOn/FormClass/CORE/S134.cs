using System;
using SAPbouiCOM;

namespace PSH_BOne_AddOn.Core
{
    /// <summary>
    /// BP마스터
    /// </summary>
    internal class S134 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private string CardCode;//거래처코드
        private string FrDate;//비활성 시작일
        private string ToDate; //비활정 종료일
        private string CreditLineV;//여신한도
        private string DflAccountV; //계좌번호
        private string BankCodeV; //은행코드
        private string AcctNameV; //은행계좌이름
        private string DflBranch; //지점명
        private SAPbouiCOM.BoFormMode oFormMode01; //클래스에서 선택한 마지막 아이템 Uid값

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
                SubMain.Add_Forms(this, formUID, "S134");

                PS_S134_CreateItems();
                PS_S134_FormItemEnabled();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private void PS_S134_CreateItems()
        {
            SAPbouiCOM.Item oNewITEM = null;

            try
            {
                oForm.Freeze(true);
                oNewITEM = oForm.Items.Add("Text", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Top = oForm.Items.Item("10002046").Top + 23;
                oNewITEM.Left = oForm.Items.Item("10002046").Left + 20;
                oNewITEM.Height = oForm.Items.Item("10002046").Height;
                oNewITEM.Width = 80;
                oNewITEM.Specific.Caption = "코드사용여부";

                oNewITEM = oForm.Items.Add("CheckYN", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Left = 120;
                oNewITEM.Top = oForm.Items.Item("10002046").Top + 23;
                oNewITEM.Height = oForm.Items.Item("10002046").Height;
                oNewITEM.Width = 80;
                oNewITEM.DisplayDesc = true;

                oNewITEM.Specific.ValidValues.Add("", "-");
                oNewITEM.Specific.ValidValues.Add("Y", "사용");
                oNewITEM.Specific.ValidValues.Add("N", "미사용");

                oNewITEM = oForm.Items.Add("Managed", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                oNewITEM.Top = oForm.Items.Item("10002046").Top + 23;
                oNewITEM.Left = 220;
                oNewITEM.Height = oForm.Items.Item("10002046").Height;
                oNewITEM.Width = 120;
                oNewITEM.Specific.Caption = "채권관리업체";
                oNewITEM.Specific.DataBind.SetBound(true, "OCRD", "U_Managed");


                oNewITEM = oForm.Items.Add("CreditLn", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oNewITEM.Left = 351;
                oNewITEM.Top = oForm.Items.Item("10002046").Top + 23;
                oNewITEM.Height = oForm.Items.Item("10002046").Height;
                oNewITEM.Width = 80;
                oNewITEM.Visible = false;
                oNewITEM.Specific.DataBind.SetBound(true, "OCRD", "U_CreditLn");

                oNewITEM = oForm.Items.Add("DflAcct", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oNewITEM.Left = 350;
                oNewITEM.Top = oForm.Items.Item("10002046").Top + 23;
                oNewITEM.Height = oForm.Items.Item("10002046").Height;
                oNewITEM.Width = 80;
                oNewITEM.Visible = false;
                oNewITEM.Specific.DataBind.SetBound(true, "OCRD", "U_DflAcct");

                oNewITEM = oForm.Items.Add("Text1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Top = oForm.Items.Item("1").Top - 12;
                oNewITEM.Left = oForm.Items.Item("1").Left;
                oNewITEM.Height = 12;
                oNewITEM.Width = 120;
                oNewITEM.FontSize = 10;
                oNewITEM.Specific.Caption = "Addon running...";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oNewITEM);
            }
        }

        /// <summary>
        /// 각 모드에 따른 아이템설정
        /// </summary>
        private void PS_S134_FormItemEnabled()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                CreditLineV = oForm.Items.Item("85").Specific.Value.ToString().Trim(); //여신한도
                DflAccountV = oForm.Items.Item("89").Specific.Value.ToString().Trim(); //계좌번호
                BankCodeV =oForm.Items.Item("434").Specific.Value.ToString().Trim(); //은행코드
                AcctNameV =oForm.Items.Item("436").Specific.Value.ToString().Trim(); //은행계좌이름
                DflBranch =oForm.Items.Item("119").Specific.Value.ToString().Trim(); //지점명

                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    sQry = "select U_Module from [@PS_SY005L]  where Code = 'S134' and U_USERID = '" + PSH_Globals.oCompany.UserName + "'";
                    oRecordSet.DoQuery(sQry);
                   
                    if (oRecordSet.Fields.Item(0).Value == "M2")  //고객
                    {
                        oForm.Items.Item("40").Specific.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue); //공급업체
                    }
                    else if (oRecordSet.Fields.Item(0).Value == "M3")
                    {
                        oForm.Items.Item("40").Specific.Select("S", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                }
                oForm.Items.Item("10002051").Enabled = false;
                oForm.Items.Item("10002054").Enabled = false;
                oForm.Items.Item("10002044").Enabled = false;
                oForm.Items.Item("10002045").Enabled = false;
                oForm.Items.Item("10002046").Enabled = false;
                CreditLineV = oForm.Items.Item("85").Specific.Value;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_S134_Power_Check
        /// </summary>
        /// <returns></returns>
        private bool PS_S134_Power_Check()
        {
            string errMessage = string.Empty;
            bool returnValue = false; 
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "select U_Power from [@PS_SY005L] where 1=1 and code ='S134' and U_USERID ='" + PSH_Globals.oCompany.UserName + "'";
                oRecordSet.DoQuery(sQry);

                if (string.IsNullOrEmpty(oRecordSet.Fields.Item(0).Value))
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return returnValue;
        }

        /// <summary>
        /// PS_S134_Credit_Check
        /// </summary>
        /// <returns></returns>
        private bool PS_S134_Credit_Check()
        {
            string errMessage = string.Empty;
            bool returnValue = false; 
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "SELECT U_Module FROM [@PS_SY005L] WHERE Code = 'S134' AND U_USERID = '" + PSH_Globals.oCompany.UserName + "'";

                oRecordSet.DoQuery(sQry);
                
                if (CreditLineV.ToString().Trim() != oForm.Items.Item("85").Specific.Value.ToString().Trim())
                {
                    sQry = "select count(*) from [@PS_SY005L]  where U_power = 'A1' and Code = 'OCRD' and U_USERID = '" + PSH_Globals.oCompany.UserName + "'";
                    oRecordSet.DoQuery(sQry);
                    if (oRecordSet.Fields.Item(0).Value > 0)
                    {
                        oForm.Items.Item("CreditLn").Specific.Value = "Y";
                    }
                    else
                    {
                        errMessage = "여신한도 수정 권한을 가지고 있지 않습니다.";
                        throw new Exception();
                    }
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return returnValue;
        }

        /// <summary>
        /// PS_S134_Account_Check
        /// </summary>
        /// <returns></returns>
        private bool PS_S134_Account_Check()
        {
            string errMessage = string.Empty;
            bool returnValue = false;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "SELECT U_Module FROM [@PS_SY005L] WHERE Code = 'S134' AND U_USERID = '" + PSH_Globals.oCompany.UserName + "'";

                oRecordSet.DoQuery(sQry);

                if (DflAccountV.ToString().Trim() != oForm.Items.Item("89").Specific.Value.ToString().Trim() || BankCodeV.ToString().Trim() != oForm.Items.Item("434").Specific.Value.ToString().Trim() || AcctNameV.ToString().Trim() != oForm.Items.Item("436").Specific.Value.ToString().Trim() || DflBranch.ToString().Trim() != oForm.Items.Item("119").Specific.Value.ToString().Trim())
                {
                    sQry = "select count(*) from [@PS_SY005L]  where U_power = 'A2' and Code = 'OCRD' and U_USERID = '" + PSH_Globals.oCompany.UserName + "'";
                    oRecordSet.DoQuery(sQry);
                    if (oRecordSet.Fields.Item(0).Value > 0)
                    {
                        oForm.Items.Item("DflAcct").Specific.Value = "Y";
                    }
                    else
                    {
                        errMessage = "계좌번호 수정 권한을 가지고 있지 않습니다.";
                        throw new Exception();
                    }
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return returnValue;
        }

        /// <summary>
        /// PS_S134_Authority_Check
        /// </summary>
        /// <returns></returns>
        private bool PS_S134_Authority_Check()
        {
            string errMessage = string.Empty;
            bool returnValue = false;
            string selectedValue; //BP마스터의 CardType(고객/공급업체)
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                selectedValue = oForm.Items.Item("40").Specific.Selected.Value;
                sQry = "SELECT U_Module FROM [@PS_SY005L] WHERE Code = 'S134' AND U_USERID = '" + PSH_Globals.oCompany.UserName + "'";

                oRecordSet.DoQuery(sQry);

                //고객을 선택했을 때
                if (selectedValue == "C")//공급업체 
                {
                    if (oRecordSet.Fields.Item(0).Value == "M3")
                    {
                        PSH_Globals.SBO_Application.MessageBox("현재 보유한 권한은 해당 기능을 수행할 수 없습니다.");
                        throw new Exception();
                    }
                }
                else if (selectedValue == "S")//공급업체를 선택했을 때
                {
                    if (oRecordSet.Fields.Item(0).Value == "M2") //고객
                    {
                        errMessage = "현재 보유한 권한은 해당 기능을 수행할 수 없습니다.";
                        throw new Exception();
                    }
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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
                //    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
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
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_S134_Authority_Check() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_S134_Power_Check() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_S134_Credit_Check() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_S134_Account_Check() == false)
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

                            CardCode = oForm.Items.Item("5").Specific.Value;

                            oFormMode01 = oForm.Mode;
                            oForm.Items.Item("CheckYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {

                            if (PS_S134_Authority_Check() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_S134_Power_Check() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_S134_Credit_Check() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_S134_Account_Check() == false)
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

                            CardCode = oForm.Items.Item("5").Specific.Value;

                            oFormMode01 = oForm.Mode;
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
                                PS_S134_FormItemEnabled();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_S134_FormItemEnabled();
                            }
                        }
                        if (pVal.ActionSuccess == true)
                        {
                            if (oForm.Items.Item("10002045").Specific.Selected == false)
                            {
                                if (oFormMode01 == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) // 업데이트시 마지막에 업데이트 일자 처리
                                {
                                    sQry = "update OCRD set UpdateDate='" + FrDate + "', validFor = 'N', frozenFor    = 'Y', frozenFrom = '" + FrDate + "', frozenTo = '" + ToDate + "', FrozenComm = '업데이트됨' FROM OCRD WHERE CardCode ='" + CardCode + "'";
                                    oRecordSet.DoQuery(sQry);
                                    oFormMode01 = SAPbouiCOM.BoFormMode.fm_OK_MODE;

                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                    oForm.Items.Item("5").Specific.Value = CardCode;
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                                }
                                else if (oFormMode01 == SAPbouiCOM.BoFormMode.fm_ADD_MODE) // 신규추가시 마지막에 업데이트 일자 처리
                                {
                                    sQry = "update OCRD set UpdateDate='" + FrDate + "', validFor = 'N', frozenFor    = 'Y', frozenFrom = '" + FrDate + "', frozenTo = '" + ToDate + "', FrozenComm = '초기등록' FROM OCRD WHERE CardCode ='" + CardCode + "'";
                                    oRecordSet.DoQuery(sQry);
                                    oFormMode01 = SAPbouiCOM.BoFormMode.fm_OK_MODE;

                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                    oForm.Items.Item("5").Specific.Value = CardCode;
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
            short errCode = 0;
            string selectedValue = null; //BP마스터의 CardType(고객/공급업체)

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {

                }
                else if (pVal.BeforeAction == false)
                {

                    if (pVal.ItemUID == "40")
                    {
                        selectedValue = oForm.Items.Item("40").Specific.Selected.Value;

                        if (PS_S134_Authority_Check() == false)
                        {
                            errCode = 1;
                            throw new Exception();
                        }

                    }
                    if (oForm.Items.Item("CheckYN").Specific.Value.ToString().Trim() == "Y")
                    {
                        oForm.Items.Item("10002044").Enabled = true;
                        oForm.Items.Item("10002045").Enabled = true;
                        oForm.Items.Item("10002046").Enabled = true;
                        oForm.Items.Item("10002044").Specific.Selected = true;

                    }
                    if (oForm.Items.Item("CheckYN").Specific.Value.ToString().Trim() == "N")
                    {
                        oForm.Items.Item("10002044").Enabled = true;
                        oForm.Items.Item("10002045").Enabled = true;
                        oForm.Items.Item("10002046").Enabled = true;
                        oForm.Items.Item("10002045").Specific.Selected = true;
                    }

                }
                oForm.Items.Item("10002051").Enabled = false;
                oForm.Items.Item("10002054").Enabled = false;
                oForm.Items.Item("10002044").Enabled = false;
                oForm.Items.Item("10002045").Enabled = false;
                oForm.Items.Item("10002046").Enabled = false;
            }
            catch (Exception ex)
            {
                if (errCode == 1) //고객을 선택한 경우는 공급업체로 강제 선택
                {
                    if (selectedValue == "C")
                    {
                        oForm.Items.Item("40").Specific.Select("S", SAPbouiCOM.BoSearchKey.psk_ByValue); //공급업체를 선택한 경우는 고객으로 강제 선택
                    }
                    else if (selectedValue == "S")
                    {
                        oForm.Items.Item("40").Specific.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
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
            try
            {
                oForm.Items.Item("10002044").Enabled = false; //활성
                oForm.Items.Item("10002045").Enabled = false; //비활성
                oForm.Items.Item("10002046").Enabled = false; //고급
                oForm.Items.Item("10002047").Enabled = false; //고급
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// MATRIX_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PSH_Globals.SBO_Application.MessageBox(oForm.Items.Item("40").Specific.Selected.Value);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "1287": //복제
                            if (PS_S134_Authority_Check() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PS_S134_FormItemEnabled();
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
                            PS_S134_FormItemEnabled();
                            break;
                        case "1282": //추가
                            PS_S134_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_S134_FormItemEnabled();
                            break;
                    }
                }
            }
            catch (Exception ex)
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
