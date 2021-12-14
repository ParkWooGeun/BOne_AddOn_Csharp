using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn.Core
{
	/// <summary>
	/// AR대변메모
	/// </summary>
	internal class S179 : PSH_BaseClass
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
				SubMain.Add_Forms(this, formUID, "S179");

                PS_S179_CreateItems();
                PS_S179_EnableFormItem(false);
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
        private void PS_S179_CreateItems()
        {
            SAPbouiCOM.Item oNewITEM = null;
            
            try
            {
                oNewITEM = oForm.Items.Add("TradeType", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Left = oForm.Items.Item("2003").Left;
                oNewITEM.Top = oForm.Items.Item("2003").Top + oForm.Items.Item("2003").Height + 1;
                oNewITEM.Height = oForm.Items.Item("2003").Height;
                oNewITEM.Width = oForm.Items.Item("2003").Width;
                oNewITEM.DisplayDesc = true;
                oNewITEM.Specific.DataBind.SetBound(true, "ORIN", "U_TradeType");
                oNewITEM.Specific.ValidValues.Add("1", "일반");
                oNewITEM.Specific.ValidValues.Add("2", "임가공");

                oNewITEM = oForm.Items.Add("Static01", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Left = oForm.Items.Item("2002").Left;
                oNewITEM.Top = oForm.Items.Item("2002").Top + oForm.Items.Item("2002").Height + 1;
                oNewITEM.Height = oForm.Items.Item("2002").Height;
                oNewITEM.Width = oForm.Items.Item("2002").Width;
                oNewITEM.LinkTo = "TradeType";
                oNewITEM.Specific.Caption = "거래형태";

                oNewITEM = oForm.Items.Add("DCardCod", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oNewITEM.Left = oForm.Items.Item("222").Left;
                oNewITEM.Top = oForm.Items.Item("222").Top + oForm.Items.Item("222").Height + 1;
                oNewITEM.Height = oForm.Items.Item("222").Height;
                oNewITEM.Width = oForm.Items.Item("222").Width;
                oNewITEM.Specific.DataBind.SetBound(true, "ORIN", "U_DCardCod");

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
                oNewITEM.Specific.DataBind.SetBound(true, "ORIN", "U_DCardNam");

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
                oNewITEM.Specific.DataBind.SetBound(true, "ORIN", "U_LotNo");

                oNewITEM = oForm.Items.Add("Static05", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Left = oForm.Items.Item("Static04").Left;
                oNewITEM.Top = oForm.Items.Item("Static04").Top + oForm.Items.Item("Static04").Height + 1;
                oNewITEM.Height = oForm.Items.Item("Static04").Height;
                oNewITEM.Width = oForm.Items.Item("Static04").Width;
                oNewITEM.LinkTo = "LotNo";
                oNewITEM.Specific.Caption = "업체수주번호";

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
        private void PS_S179_EnableFormItem(bool Status)
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
        private bool PS_S179_CheckDataValid()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
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
                else if (string.IsNullOrEmpty(oForm.Items.Item("2001").Specific.Value.ToString()))
                {
                    errMessage = "사업장은 필수입니다.";
                    oForm.Items.Item("2001").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("TradeType").Specific.Value.ToString().Trim()))
                {
                    errMessage = "거래형태는 필수입니다.";
                    oForm.Items.Item("TradeType").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (oForm.Items.Item("2001").Specific.Value.ToString().Trim() != "1" && oForm.Items.Item("TradeType").Specific.Selected.Value == "2") //창원이 아닌경우 임가공 선택한 경우
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
                        returnValue = false;
                        return returnValue;
                    }
                    
                    if (oForm.Items.Item("70").Specific.Selected.Value == "S" | oForm.Items.Item("70").Specific.Selected.Value == "L") //현지,시스템통화
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
                        if (dataHelpClass.GetItem_TradeType(oMat01.Columns.Item("1").Cells.Item(i).Specific.Value) == "1") //품목 : 일반
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



        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	//BeforeAction = True
        //	if ((pVal.BeforeAction == true)) {
        //		switch (pVal.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
        //				break;
        //			case "1281":
        //				//찾기
        //				break;
        //			case "1282":
        //				//추가
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				break;
        //		}
        //	//BeforeAction = False
        //	} else if ((pVal.BeforeAction == false)) {
        //		switch (pVal.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
        //				break;
        //			case "1281":
        //				//찾기
        //				PS_S179_EnableFormItem();
        //				break;
        //			case "1282":
        //				//추가
        //				PS_S179_EnableFormItem();
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				PS_S179_EnableFormItem();
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_MenuEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	//BeforeAction = True
        //	if ((BusinessObjectInfo.BeforeAction == true)) {
        //		switch (BusinessObjectInfo.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				//33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				//34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				//35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				//36
        //				break;
        //		}
        //	//BeforeAction = False
        //	} else if ((BusinessObjectInfo.BeforeAction == false)) {
        //		switch (BusinessObjectInfo.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				//33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				//34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				//35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				//36
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_FormDataEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //		//        If pVal.ItemUID = "Mat01" And pVal.Row > 0 And pVal.Row <= oMat01.RowCount Then
        //		//            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
        //		//            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        //		//            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
        //		//            MenuCreationParams01.uniqueID = "MenuUID"
        //		//            MenuCreationParams01.String = "메뉴명"
        //		//            MenuCreationParams01.Enabled = True
        //		//            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
        //		//        End If
        //	} else if (pVal.BeforeAction == false) {
        //		//        If pVal.ItemUID = "Mat01" And pVal.Row > 0 And pVal.Row <= oMat01.RowCount Then
        //		//                Call Sbo_Application.Menus.RemoveEx("MenuUID")
        //		//        End If
        //	}
        //	if (pVal.ItemUID == "38") {
        //		if (pVal.Row > 0) {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = pVal.ColUID;
        //			oLastColRow01 = pVal.Row;
        //		}
        //	} else {
        //		oLastItemUID01 = pVal.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}
        //	return;
        //	Raise_RightClickEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ITEM_PRESSED
        //private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemUID == "1") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (PS_S179_CheckDataValid() == false) {
        //					BubbleEvent = false;
        //					return;
        //				}
        //				//해야할일 작업
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //				if (PS_S179_CheckDataValid() == false) {
        //					BubbleEvent = false;
        //					return;
        //				}
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		if (pVal.ItemUID == "Button01") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				//UPGRADE_WARNING: oForm.Items(Combo01).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (oForm.Items.Item("Combo01").Specific.Selected == null) {
        //				} else {
        //					PS_S179_Print_Report01();
        //				}
        //			}
        //		}
        //	} else if (pVal.BeforeAction == false) {
        //		if (pVal.ItemUID == "1") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (pVal.ActionSuccess == true) {
        //					PS_S179_EnableFormItem();
        //				}
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				if (pVal.ActionSuccess == true) {
        //					PS_S179_EnableFormItem();
        //				}
        //			}
        //		}
        //	}
        //	return;
        //	Raise_EVENT_ITEM_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_KEY_DOWN
        //private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string TradeType = null;
        //	object ChildForm01 = null;
        //	if (pVal.BeforeAction == true) {
        //		//        Call dataHelpClass.ActiveUserDefineValue(oForm, pVal, BubbleEvent, "ItemCode", "") '//사용자값활성
        //		//        Call dataHelpClass.ActiveUserDefineValue(oForm, pVal, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
        //		dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "DCardCod", "");
        //		//사용자값활성
        //		if ((pVal.ItemUID == "38")) {
        //			//품목코드 변경시
        //			if ((pVal.ColUID == "1")) {
        //				if (pVal.CharPressed == 9) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					TradeType = Strings.Trim(oForm.Items.Item("TradeType").Specific.Selected.Value);

        //					ChildForm01 = new PS_SM020();
        //					//UPGRADE_WARNING: ChildForm01.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					ChildForm01.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, oMat01.VisualRowCount, TradeType);
        //					BubbleEvent = false;
        //					return;
        //				}
        //			}
        //		}
        //		//        Call dataHelpClass.ActiveUserDefineValue(oForm, pVal, BubbleEvent, "38", "U_SD030Num") '//사용자값활성
        //		//        Call dataHelpClass.ActiveUserDefineValueAlways(oForm, pVal, BubbleEvent, "38", "U_Unweight")
        //		//        Call dataHelpClass.ActiveUserDefineValueAlways_Price(oForm, pVal, BubbleEvent, "38", "14")
        //		//        Call dataHelpClass.ActiveUserDefineValueAlways_UnitWeight(oForm, pVal, BubbleEvent, "38", "11")
        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_KEY_DOWN_Error:
        //	if (Err().Number == Convert.ToDouble("-7008")) {
        //		MDC_Com.MDC_GF_Message(ref "사용자정의필드가 활성화되어 있지 않습니다.", ref "W");
        //		BubbleEvent = false;
        //		return;
        //	}
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_COMBO_SELECT
        //private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		SubMain.Sbo_Application.Forms.GetForm("-" + oForm.Type, oForm.TypeCount).Update();
        //		if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			PS_S179_EnableFormItem(true);
        //		}
        //	}
        //	oForm.Freeze(false);
        //	return;
        //	Raise_EVENT_COMBO_SELECT_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CLICK
        //private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		if ((pVal.ItemUID == "10000330")) {
        //			if (pVal.ActionSuccess == true) {
        //				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //					oSetBackOrderFunction01 = true;
        //				}
        //			}
        //		}
        //	}
        //	return;
        //	Raise_EVENT_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_VALIDATE
        //private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	string itemCode = null;
        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemChanged == true) {
        //			//매트릭스
        //			if ((pVal.ItemUID == "38")) {
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				itemCode = oMat01.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value;
        //				//수량필드 값변경시
        //				if ((pVal.ColUID == "U_Qty")) {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if ((Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) <= 0)) {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value = 0;
        //						//수량
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
        //						//중량
        //					} else {
        //						//EA자체품
        //						if ((dataHelpClass.GetItem_SbasUnit(itemCode) == "101")) {
        //							//UPGRADE_WARNING: oMat01.Columns(11).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value);
        //						//EAUOM
        //						} else if ((dataHelpClass.GetItem_SbasUnit(itemCode) == "102")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if (Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(itemCode)) == 0) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
        //							} else {
        //								//UPGRADE_WARNING: oMat01.Columns(11).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(itemCode));
        //							}
        //						//KGSPEC
        //						} else if ((dataHelpClass.GetItem_SbasUnit(itemCode) == "201")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if ((Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) == 0) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
        //							} else {
        //								//UPGRADE_WARNING: oMat01.Columns(11).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = (Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value);
        //							}
        //						//KG단중
        //						} else if ((dataHelpClass.GetItem_SbasUnit(itemCode) == "202")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if (System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0) == 0) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
        //							} else {
        //								//UPGRADE_WARNING: oMat01.Columns(11).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0);
        //							}
        //						//KG입력
        //						} else if ((dataHelpClass.GetItem_SbasUnit(itemCode) == "203")) {
        //						}
        //					}
        //				} else if ((pVal.ColUID == "11")) {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if ((Convert.ToDouble(oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value) <= 0)) {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value = 0;
        //						//수량
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
        //						//중량
        //					} else {
        //						//EA자체품
        //						if ((dataHelpClass.GetItem_SbasUnit(itemCode) == "101")) {
        //						//EAUOM
        //						} else if ((dataHelpClass.GetItem_SbasUnit(itemCode) == "102")) {
        //						//KGSPEC
        //						} else if ((dataHelpClass.GetItem_SbasUnit(itemCode) == "201")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if ((Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) == 0) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
        //							} else {
        //								//UPGRADE_WARNING: oMat01.Columns(11).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = (Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value);
        //							}
        //						//KG단중
        //						} else if ((dataHelpClass.GetItem_SbasUnit(itemCode) == "202")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if (System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0) == 0) {
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
        //							} else {
        //								//UPGRADE_WARNING: oMat01.Columns(11).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0);
        //							}
        //						//KG입력
        //						} else if ((dataHelpClass.GetItem_SbasUnit(itemCode) == "203")) {
        //						}
        //					}
        //				} else if (pVal.ColUID == "1") {
        //					if (oMat01.VisualRowCount > 1) {
        //						oForm.Items.Item("TradeType").Enabled = false;
        //					} else {
        //						oForm.Items.Item("TradeType").Enabled = true;
        //					}
        //				}
        //			} else if (pVal.ItemUID == "DCardCod") {
        //				//UPGRADE_WARNING: oForm.Items(DCardNam).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: dataHelpClass.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("DCardNam").Specific.Value = dataHelpClass.GetValue("SELECT CardName FROM OCRD WHERE CardCode = '" + oForm.Items.Item("DCardCod").Specific.Value + "'", 0, 1);
        //			}
        //		}
        //		SubMain.Sbo_Application.Forms.GetForm("-" + oForm.Type, oForm.TypeCount).Update();
        //	} else if (pVal.BeforeAction == false) {
        //		PS_S179_EnableFormItem(true);
        //	}
        //	oForm.Freeze(false);
        //	return;
        //	Raise_EVENT_VALIDATE_Error:
        //	oForm.Freeze(false);
        //	if (Err().Number == Convert.ToDouble("-7008 ")) {
        //	} else {
        //		SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	}
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LOAD
        //private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_MATRIX_LOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_GOT_FOCUS
        //private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.ItemUID == "38") {
        //		if (pVal.Row > 0) {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = pVal.ColUID;
        //			oLastColRow01 = pVal.Row;
        //		}
        //	} else {
        //		oLastItemUID01 = pVal.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}
        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		if ((oSetBackOrderFunction01 == true)) {
        //			oSetBackOrderFunction01 = false;
        //			dataHelpClass.SBO_SetBackOrderFunction(ref oForm);
        //		}
        //	}
        //	return;
        //	Raise_EVENT_GOT_FOCUS_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_FORM_UNLOAD
        //private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //	} else if (pVal.BeforeAction == false) {
        //		SubMain.RemoveForms(oFormUniqueID);
        //		//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oForm = null;
        //		//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oMat01 = null;
        //	}
        //	return;
        //	Raise_EVENT_FORM_UNLOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ROW_DELETE
        //private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	if ((oLastColRow01 > 0)) {
        //		if (pVal.BeforeAction == true) {
        //			//행삭제전 행삭제가능여부검사
        //		} else if (pVal.BeforeAction == false) {
        //			if (oMat01.VisualRowCount > 1) {
        //				oForm.Items.Item("TradeType").Enabled = false;
        //			} else {
        //				oForm.Items.Item("TradeType").Enabled = true;
        //			}
        //			//        For i = 1 To oMat01.VisualRowCount
        //			//            oMat01.Columns("COL01").Cells(i).Specific.Value = i
        //			//        Next i
        //			//        oMat01.FlushToDataSource
        //			//        Call oDS_ZYM30L.RemoveRecord(oDS_ZYM30L.Size - 1)
        //			//        oMat01.LoadFromDataSource
        //			//        If oMat01.RowCount = 0 Then
        //			//            Call PS_SD380_AddMatrixRow(0)
        //			//        Else
        //			//            If Trim(oDS_SD380L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
        //			//                Call PS_SD380_AddMatrixRow(oMat01.RowCount)
        //			//            End If
        //			//        End If
        //		}
        //	}
        //	return;
        //	Raise_EVENT_ROW_DELETE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion
    }
}
