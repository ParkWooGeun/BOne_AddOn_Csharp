using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using System.Collections.Generic;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn.Core
{
	/// <summary>
	/// 판매견적
	/// </summary>
	internal class S149 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private bool oSetBackOrderFunction01;

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
				oMat01 = oForm.Items.Item("38").Specific;
				SubMain.Add_Forms(this, formUID, "S149");

				S149_CreateItems();
				S149_EnableFormItem(false);
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
        private void S149_CreateItems()
        {
            SAPbouiCOM.Item oNewITEM = null;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oNewITEM = oForm.Items.Add("TradeType", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Left = oForm.Items.Item("2003").Left;
                oNewITEM.Top = oForm.Items.Item("2003").Top + oForm.Items.Item("2003").Height + 1;
                oNewITEM.Height = oForm.Items.Item("2003").Height;
                oNewITEM.Width = oForm.Items.Item("2003").Width;
                oNewITEM.DisplayDesc = true;
                oNewITEM.Specific.DataBind.SetBound(true, "OQUT", "U_TradeType");
                oNewITEM.Specific.ValidValues.Add("1", "일반");
                oNewITEM.Specific.ValidValues.Add("2", "임가공");

                oNewITEM = oForm.Items.Add("Static01", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Left = oForm.Items.Item("2002").Left;
                oNewITEM.Top = oForm.Items.Item("2002").Top + oForm.Items.Item("2002").Height + 1;
                oNewITEM.Height = oForm.Items.Item("2002").Height;
                oNewITEM.Width = oForm.Items.Item("2002").Width;
                oNewITEM.LinkTo = "TradeType";
                oNewITEM.Specific.Caption = "거래형태";

                oNewITEM = oForm.Items.Add("BPLId", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Left = oForm.Items.Item("14").Left;
                oNewITEM.Top = oForm.Items.Item("14").Top + oForm.Items.Item("14").Height * 2 + 1;
                oNewITEM.Height = oForm.Items.Item("14").Height;
                oNewITEM.Width = oForm.Items.Item("14").Width;
                oNewITEM.DisplayDesc = true;
                oNewITEM.Specific.DataBind.SetBound(true, "OQUT", "U_BPLId");
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", dataHelpClass.User_BPLID(), false, false);
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                oNewITEM = oForm.Items.Add("Static02", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Left = oForm.Items.Item("70").Left;
                oNewITEM.Top = oForm.Items.Item("70").Top + oForm.Items.Item("70").Height + 1;
                oNewITEM.Height = oForm.Items.Item("70").Height;
                oNewITEM.Width = oForm.Items.Item("70").Width;
                oNewITEM.LinkTo = "BPLId";
                oNewITEM.Specific.Caption = "사업장";

                oNewITEM = oForm.Items.Add("DCardCod", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oNewITEM.Left = oForm.Items.Item("222").Left;
                oNewITEM.Top = oForm.Items.Item("222").Top + oForm.Items.Item("222").Height + 1;
                oNewITEM.Height = oForm.Items.Item("222").Height;
                oNewITEM.Width = oForm.Items.Item("222").Width;
                oNewITEM.Specific.DataBind.SetBound(true, "OQUT", "U_DCardCod");

                oNewITEM = oForm.Items.Add("Static03", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Left = oForm.Items.Item("230").Left;
                oNewITEM.Top = oForm.Items.Item("230").Top + oForm.Items.Item("230").Height + 1;
                oNewITEM.Height = oForm.Items.Item("230").Height;
                oNewITEM.Width = oForm.Items.Item("230").Width;
                oNewITEM.LinkTo = "DCardCod";
                oNewITEM.Specific.Caption = "납품처코드";

                oNewITEM = oForm.Items.Add("DCardNam", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oNewITEM.Left = oForm.Items.Item("DCardCod").Left;
                oNewITEM.Top = oForm.Items.Item("DCardCod").Top + oForm.Items.Item("DCardCod").Height + 1;
                oNewITEM.Height = oForm.Items.Item("DCardCod").Height;
                oNewITEM.Width = oForm.Items.Item("DCardCod").Width;
                oNewITEM.Enabled = false;
                oNewITEM.Specific.DataBind.SetBound(true, "OQUT", "U_DCardNam");

                oNewITEM = oForm.Items.Add("Static04", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Left = oForm.Items.Item("Static03").Left;
                oNewITEM.Top = oForm.Items.Item("Static03").Top + oForm.Items.Item("Static03").Height + 1;
                oNewITEM.Height = oForm.Items.Item("Static03").Height;
                oNewITEM.Width = oForm.Items.Item("Static03").Width;
                oNewITEM.LinkTo = "DCardNam";
                oNewITEM.Specific.Caption = "납품처명";

                oNewITEM = oForm.Items.Add("LotNo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oNewITEM.Left = oForm.Items.Item("DCardNam").Left;
                oNewITEM.Top = oForm.Items.Item("DCardNam").Top + oForm.Items.Item("DCardNam").Height + 1;
                oNewITEM.Height = oForm.Items.Item("DCardNam").Height;
                oNewITEM.Width = oForm.Items.Item("DCardNam").Width;
                oNewITEM.Specific.DataBind.SetBound(true, "OQUT", "U_LotNo");

                oNewITEM = oForm.Items.Add("Static05", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Left = oForm.Items.Item("Static04").Left;
                oNewITEM.Top = oForm.Items.Item("Static04").Top + oForm.Items.Item("Static04").Height + 1;
                oNewITEM.Height = oForm.Items.Item("Static04").Height;
                oNewITEM.Width = oForm.Items.Item("Static04").Width;
                oNewITEM.LinkTo = "LotNo";
                oNewITEM.Specific.Caption = "업체수주번호";

                oNewITEM = oForm.Items.Add("Button01", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oNewITEM.AffectsFormMode = false;
                oNewITEM.Left = oForm.Items.Item("10000329").Left;
                oNewITEM.Top = oForm.Items.Item("10000329").Top - 23;
                oNewITEM.Height = oForm.Items.Item("10000329").Height;
                oNewITEM.Width = oForm.Items.Item("10000329").Width;
                oNewITEM.LinkTo = "Combo01";
                oNewITEM.Specific.Caption = "출력";
                oNewITEM.Specific.Type = SAPbouiCOM.BoButtonTypes.bt_Caption;

                oNewITEM = oForm.Items.Add("Combo01", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.AffectsFormMode = false;
                oNewITEM.Left = oForm.Items.Item("10000330").Left;
                oNewITEM.Top = oForm.Items.Item("10000330").Top - 21;
                oNewITEM.Height = 16;
                oNewITEM.Width = oForm.Items.Item("10000330").Width;
                oNewITEM.LinkTo = "Button01";
                oNewITEM.DisplayDesc = true;
                dataHelpClass.Set_ComboList(oForm.Items.Item("Combo01").Specific, "SELECT U_Minor,U_CdName FROM [@PS_SY001L] WHERE CODE = 'RPT01' AND U_RelCd = '149'", "01", false, false);

                oNewITEM = oForm.Items.Add("AddonText", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Top = oForm.Items.Item("1").Top - 12;
                oNewITEM.Left = oForm.Items.Item("1").Left;
                oNewITEM.Height = 12;
                oNewITEM.Width = 120;
                oNewITEM.FontSize = 10;
                oNewITEM.Specific.Caption = "Addon running";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oNewITEM);
            }
        }

        /// <summary>
        /// 각 모드에 따른 아이템설정
        /// </summary>
        /// <param name="Status"></param>
        private void S149_EnableFormItem(bool Status)
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("TradeType").Enabled = true;
                    oForm.Items.Item("DCardNam").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("TradeType").Enabled = true;
                    oForm.Items.Item("DCardNam").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("TradeType").Enabled = false;
                    oForm.Items.Item("DCardNam").Enabled = false;
                }

                if (Status == true)
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        oForm.Items.Item("TradeType").Enabled = false;
                        oForm.Items.Item("DCardNam").Enabled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool S149_CheckDataValid()
        {
            string errMessage = string.Empty;
            bool returnValue = false;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("4").Specific.Value))
                {
                    errMessage = "고객은 필수입니다.";
                    oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (oForm.Items.Item("BPLId").Specific.Selected == null)
                {
                    errMessage = "사업장은 필수입니다.";
                    oForm.Items.Item("2001").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (oForm.Items.Item("TradeType").Specific == null)
                {
                    errMessage = "거래형태는 필수입니다.";
                    oForm.Items.Item("TradeType").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (oForm.Items.Item("BPLId").Specific.Selected.Value != "1" & oForm.Items.Item("TradeType").Specific.Selected.Value == "2") //창원이 아닌경우 임가공 선택한 경우
                {
                    errMessage = "창원사업장이 아닌경우 임가공거래가 불가능합니다.";
                    throw new Exception();
                }

                for (int i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("1").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "품목은 필수입니다.";
                        oMat01.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    if (Convert.ToDouble(oMat01.Columns.Item("11").Cells.Item(i).Specific.Value) <= 0)
                    {
                        errMessage = "수량(중량)은 필수입니다.";
                        oMat01.Columns.Item("11").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    if (string.IsNullOrEmpty(oMat01.Columns.Item("14").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "단가는 필수입니다.";
                        oMat01.Columns.Item("14").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    
                    if (oForm.Items.Item("70").Specific.Selected.Value == "S" || oForm.Items.Item("70").Specific.Selected.Value == "L") //현지,시스템통화
                    {
                        if (codeHelpClass.Right(oMat01.Columns.Item("14").Cells.Item(i).Specific.Value, 3) != "KRW")
                        {
                            errMessage = "헤더와 라인의 통화가 다릅니다.";
                            throw new Exception();
                        }
                    }
                    
                    if (oForm.Items.Item("70").Specific.Selected.Value == "C") //BP통화
                    {
                        if (oForm.Items.Item("63").Specific.Value != codeHelpClass.Right(oMat01.Columns.Item("14").Cells.Item(i).Specific.Value, 3)) //DocCur 과 Price의 마지막3자리 비교
                        {
                            errMessage = "헤더와 라인의 통화가 다릅니다.";
                            throw new Exception();
                        }
                    }
                    
                    if (oForm.Items.Item("TradeType").Specific.Selected.Value == "1") //일반
                    {
                        if (dataHelpClass.GetItem_TradeType(oMat01.Columns.Item("1").Cells.Item(i).Specific.Value) == "2") //품목 : 임가공
                        {
                            errMessage = "문서의 거래형태와 품목의 거래형태가 다릅니다.";
                            oMat01.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }
                    
                    if (oForm.Items.Item("TradeType").Specific.Selected.Value == "2") //임가공
                    {
                        if (dataHelpClass.GetItem_TradeType(oMat01.Columns.Item("1").Cells.Item(i).Specific.Value) == "1")  //품목 : 일반
                        {
                            errMessage = "문서의 거래형태와 품목의 거래형태가 다릅니다.";
                            oMat01.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
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

            return returnValue;
        }

        /// <summary>
        /// 리포트 출력
        /// </summary>
        [STAThread]
        private void S149_PrintReport()
        {
            string WinTitle;
            string ReportName = string.Empty;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            
            try
            {
                WinTitle = "견적서";

                if (oForm.Items.Item("Combo01").Specific.Selected.Value.ToString().Trim() == "01")
                {
                    ReportName = "PS_S149_01.rpt";
                    //프로시저 : PS_S149_01
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>
                {
                    new PSH_DataPackClass("@DocEntry", oForm.Items.Item("8").Specific.Value)
                };

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
            finally
            {
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
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
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
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (S149_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (S149_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    if (pVal.ItemUID == "Button01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (oForm.Items.Item("Combo01").Specific.Selected != null)
                            {
                                System.Threading.Thread thread = new System.Threading.Thread(S149_PrintReport);
                                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                                thread.Start();
                            }
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
                                S149_EnableFormItem(false);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                S149_EnableFormItem(false);
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
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "DCardCod", ""); //사용자값활성

                    if (pVal.ItemUID == "38")
                    {
                        if (pVal.ColUID == "1") //품목코드
                        {
                            if (pVal.CharPressed == 9)
                            {
                                PS_SM020 tempForm = new PS_SM020();
                                tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, oMat01.VisualRowCount, oForm.Items.Item("TradeType").Specific.Selected.Value.ToString().Trim());
                                BubbleEvent = false;
                                return;
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
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
        /// GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.ItemUID == "38")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = pVal.ColUID;
                        oLastColRow01 = pVal.Row;
                    }
                }
                else
                {
                    oLastItemUID01 = pVal.ItemUID;
                    oLastColUID01 = "";
                    oLastColRow01 = 0;
                }

                if (pVal.BeforeAction == true)
                {

                }
                else if (pVal.BeforeAction == false)
                {
                    if (oSetBackOrderFunction01 == true)
                    {
                        oSetBackOrderFunction01 = false;
                        dataHelpClass.SBO_SetBackOrderFunction(ref oForm);
                    }
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
                    PSH_Globals.SBO_Application.Forms.GetForm(oForm.Type.ToString(), oForm.TypeCount).Update();

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        S149_EnableFormItem(true);
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
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "10000330")
                    {
                        if (pVal.ActionSuccess == true)
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                oSetBackOrderFunction01 = true;
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
            string itemCode;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "38") //매트릭스
                        {
                            itemCode = oMat01.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value;
                            
                            if (pVal.ColUID == "U_Qty") //수량
                            {
                                if (Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) <= 0)
                                {
                                    oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value = 0; //수량
                                    oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1; //중량
                                }
                                else
                                {
                                    if (dataHelpClass.GetItem_SbasUnit(itemCode) == "101") //EA자체품
                                    {
                                        oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "102") //EAUOM
                                    {
                                        if (Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(itemCode)) == 0)
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(itemCode));
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "201") //KGSPEC
                                    {
                                        if ((Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) == 0)
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = (Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value);
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "202") //KG단중
                                    {
                                        if (System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0) == 0)
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0);
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "203") //KG입력
                                    {
                                    }
                                }
                            }
                            else if (pVal.ColUID == "11")
                            {
                                if (Convert.ToDouble(oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value) <= 0)
                                {
                                    oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value = 0; //수량
                                    oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1; //중량
                                }
                                else
                                {
                                    if (dataHelpClass.GetItem_SbasUnit(itemCode) == "101") //EA자체품
                                    {
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "102") //EAUOM
                                    {   
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "201") //KGSPEC
                                    {
                                        if ((Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) == 0)
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = (Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value);
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "202") //KG단중
                                    {
                                        if (System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0) == 0)
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0);
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "203") //KG입력
                                    {
                                    }
                                }
                            }
                            else if (pVal.ColUID == "1")
                            {
                                if (oMat01.VisualRowCount > 1)
                                {
                                    oForm.Items.Item("TradeType").Enabled = false;
                                }
                                else
                                {
                                    oForm.Items.Item("TradeType").Enabled = true;
                                }
                            }
                        }
                        else if (pVal.ItemUID == "DCardCod")
                        {
                            oForm.Items.Item("DCardNam").Specific.Value = dataHelpClass.GetValue("SELECT CardName FROM OCRD WHERE CardCode = '" + oForm.Items.Item("DCardCod").Specific.Value + "'", 0, 1);
                        }
                    }

                    PSH_Globals.SBO_Application.Forms.GetForm(oForm.Type.ToString(), oForm.TypeCount).Update();
                }
                else if (pVal.Before_Action == false)
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        S149_EnableFormItem(true);
                    }
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
                    oMat01.AutoResizeColumns();
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
        /// FORM_ACTIVATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_ACTIVATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (oSetBackOrderFunction01 == true)
                    {
                        oSetBackOrderFunction01 = false;
                        dataHelpClass.SBO_SetBackOrderFunction(ref oForm);
                    }
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
        /// 행삭제 체크 메서드(Raise_FormMenuEvent 에서 사용)
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        if (oMat01.VisualRowCount > 1)
                        {
                            oForm.Items.Item("TradeType").Enabled = false;
                        }
                        else
                        {
                            oForm.Items.Item("TradeType").Enabled = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            S149_EnableFormItem(false);
                            break;
                        case "1282": //추가
                            S149_EnableFormItem(false);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            oMat01.AutoResizeColumns();
                            S149_EnableFormItem(false);
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
                        //Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        //Raise_EVENT_FORM_DATA_ADD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        //Raise_EVENT_FORM_DATA_UPDATE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        //Raise_EVENT_FORM_DATA_DELETE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
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
        /// RightClickEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                }

                if (pVal.ItemUID == "38")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = pVal.ColUID;
                        oLastColRow01 = pVal.Row;
                    }
                }
                else
                {
                    oLastItemUID01 = pVal.ItemUID;
                    oLastColUID01 = "";
                    oLastColRow01 = 0;
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
    }
}