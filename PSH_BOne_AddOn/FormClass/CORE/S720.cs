using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn.Core
{
	/// <summary>
	/// 출고
	/// </summary>
	internal class S720 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
				oMat = oForm.Items.Item("13").Specific;
				SubMain.Add_Forms(this, formUID, "S720");

                S720_CreateItems();
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
        private void S720_CreateItems()
        {
            SAPbouiCOM.Item oItem = null;

            try
            {
                //특정사용자만 출고 아이템 중량을 가져올 수 있도록 하기 위한 UserID TextBox 추가_S (2011.10.01 송명규)
                oItem = oForm.Items.Add("UserID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Left = oForm.Items.Item("21").Left; //참조2(21) Item 기준
                oItem.Top = oForm.Items.Item("21").Top + oForm.Items.Item("21").Height + 1;
                oItem.Height = oForm.Items.Item("21").Height;
                oItem.Width = oForm.Items.Item("21").Width;
                oItem.Specific.Value = PSH_Globals.oCompany.UserName; //로그인한 사용자의 ID
                oForm.Items.Item("38").Click(SAPbouiCOM.BoCellClickType.ct_Regular); //증빙일로의 포커스 강제 이동
                oItem.Visible = false; //UserID 텍스트박스 숨김
                //특정사용자만 출고 아이템 중량을 가져올 수 있도록 하기 위한 UserID TextBox 추가_E (2011.10.01 송명규)

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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool S720_CheckDataValid()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                //마감상태 체크_S(2017.11.23 송명규 추가)
                if (dataHelpClass.Check_Finish_Status(dataHelpClass.User_BPLID(), oForm.Items.Item("9").Specific.Value, oForm.TypeEx) == false)
                {
                    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다." + (char)13 + "전기일을 확인하고, 회계부서로 문의하세요.";
                    throw new Exception();
                }
                //마감상태 체크_E(2017.11.23 송명규 추가)

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
        /// Form Item Event
        /// </summary>
        /// <param name = "FormUID" > Form UID</param>
        /// <param name = "pVal" > pVal </param >
        /// <param name = "BubbleEvent">Bubble Event</param>
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
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                            if (S720_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (S720_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
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
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "13")
                    {   
                        if (pVal.ColUID == "1") //품목코드
                        {
                            if (pVal.CharPressed == 9)
                            {
                                PS_SM020 tempForm = new PS_SM020();
                                tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, oMat.VisualRowCount, "");
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (pVal.ColUID == "U_CardCode") //거래처코드
                        {
                            if (pVal.CharPressed == 9)
                            {
                                if (string.IsNullOrEmpty(oMat.Columns.Item("U_CardCode").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string itemCode;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "13") //매트릭스
                        {
                            if (pVal.ColUID == "1") //품목코드
                            {
                                string tempQuery = "SELECT U_CdName AS [WhsCode] FROM [@PS_SY001L] WHERE Code = 'I002' AND U_Minor = '" + oForm.Items.Item("UserID").Specific.Value + "'";
                                //string tempQuery = "SELECT U_CdName AS [WhsCode] FROM [@PS_SY001L] WHERE Code = 'I002' AND U_Minor = '" + PSH_Globals.oCompany.UserName + "'";
                                oRecordSet.DoQuery(tempQuery);
                                string outWshCode = oRecordSet.Fields.Item("WhsCode").Value.ToString().Trim(); //기계사업부 원재료 불출용 창고
                                string baseWshCode = outWshCode == "" ? dataHelpClass.User_WhsCode("1") : outWshCode; //기계사업부 원재료 불출용 창고가 설정되지 않은 사용자는 기본 창고, 아니면 불출용 창고 코드로 설정;
                                oMat.Columns.Item("15").Cells.Item(pVal.Row).Specific.Value = baseWshCode; //창고
                            }
                            else if (pVal.ColUID == "U_Qty")
                            {
                                if (Convert.ToDouble(oMat.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) <= 0)
                                {
                                    oMat.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value = 0; //수량
                                    oMat.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = 1; //중량
                                }
                                else
                                {
                                    itemCode = oMat.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value;

                                    if (dataHelpClass.GetItem_SbasUnit(itemCode) == "101") //EA자체품
                                    {
                                        oMat.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value); //EAUOM
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "102")
                                    {
                                        if (Convert.ToDouble(oMat.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(itemCode)) == 0)
                                        {
                                            oMat.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(itemCode));
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "201") //KGSPEC
                                    {
                                        if ((Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) == 0)
                                        {
                                            oMat.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = (Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value);
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "202") //KG단중
                                    {
                                        if (System.Math.Round(Convert.ToDouble(oMat.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0) == 0)
                                        {
                                            oMat.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = System.Math.Round(Convert.ToDouble(oMat.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 2);
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "203") //KG입력
                                    {
                                    }
                                }
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "13")
                        {
                            if (pVal.ColUID == "U_CardCode")
                            {
                                oForm.Freeze(true);
                                sQry = "select cardname from ocrd where cardcode = '" + oMat.Columns.Item("U_CardCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                                oMat.Columns.Item("U_CardName").Editable = true;
                                oMat.Columns.Item("U_CardName").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                                oMat.Columns.Item("U_CardCode").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oMat.Columns.Item("U_CardName").Editable = false;
                                oForm.Freeze(false);
                            }
                            else if (pVal.ColUID == "9")
                            {
                                string tempQuery = "SELECT U_CdName AS [WhsCode] FROM [@PS_SY001L] WHERE Code = 'I002' AND U_Minor = '" + oForm.Items.Item("UserID").Specific.Value + "'";
                                //string tempQuery = "SELECT U_CdName AS [WhsCode] FROM [@PS_SY001L] WHERE Code = 'I002' AND U_Minor = '" + PSH_Globals.oCompany.UserName + "'";
                                oRecordSet.DoQuery(tempQuery);
                                string outWshCode = oRecordSet.Fields.Item("WhsCode").Value.ToString().Trim(); //기계사업부 원재료 불출용 창고
                                
                                if (outWshCode != "") //기계사업부 원재료 불출용 창고 설정이 되어 있는 사용자일경우
                                {
                                    //총계를 계산하지 않고, AP송장의 총계를 조회
                                    sQry = "EXEC S720_01 '" + oMat.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "', '" + oForm.Items.Item("9").Specific.Value.ToString().Trim() + "'";
                                    oRecordSet.DoQuery(sQry);
                                    oMat.Columns.Item("14").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item("LineTotal").Value.ToString().Trim();
                                }
                            }
                        }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
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
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            oMat.AutoResizeColumns();
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
                switch (pVal.ItemUID)
                {
                    case "Mat01":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
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
    }
}
